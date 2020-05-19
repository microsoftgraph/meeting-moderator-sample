require('dotenv').config({ path: '.env' });
const {
    TurnContext,
    MessageFactory,
    TeamsInfo,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');
const TextEncoder = require('util').TextEncoder;
const { CosmosClient } = require("@azure/cosmos");
const endpoint = process.env.COSMOS_ENDPOINT;
const key = process.env.COSMOS_KEY;
const databaseId = process.env.COSMOS_DATABASEID;
const containerId = process.env.COSMOS_CONTAINERID;
const cosmosClient = new CosmosClient({ endpoint, key });

class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();
            if (text.includes("mention")) {
                console.log(text);
                console.log(text.includes("QQ"));
                await this.mentionActivityAsync(context);
            } else if(text.includes("update")) {
                await this.cardActivityAsync(context, true);
            } else if (text.includes("delete")) {
                await this.deleteCardActivityAsync(context);
            } else if (text.includes("message")) {
                await this.messageAllMembersAsync(context);
            } else if (text.includes("who")) {
                await this.getSingleMember(context);
            } else if (text.includes("qq")) {
                await this.enqueueQuestion(context);
            } else if (text.includes("next question")) {
                await this.dequeueQuestion(context);
            } else {
                await this.cardActivityAsync(context, false)
            }   
        });

        this.onMembersAddedActivity(async (context, next) => {
            context.activity.membersAdded.forEach(async (teamMember) => {
                if (teamMember.id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to the team ${ teamMember.givenName } ${ teamMember.surname }`);
                }
            });
            await next();
        });
    }

    async cardActivityAsync(context, isUpdate) {
        const cardActions = [                       
                {
                    type: ActionTypes.MessageBack,
                    title: 'Message all members',
                    value: null,
                    text: 'MessageAllMembers'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Who am I?',
                    value: null,
                    text: 'whoami'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Delete card',
                    value: null,
                    text: 'Delete'
                }
        ];

        if(isUpdate) {
            await this.sendUpdateCard(context, cardActions);
        }
        else {
            await this.sendWelcomeCard(context, cardActions);
        }
    }

    async sendUpdateCard(context, cardActions) {
        const data = context.activity.value;
        data.count += 1;
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: data,
            text: 'UpdateCardAction'
        });
        const card = CardFactory.heroCard(
            'Updated card',
            `Update count: ${data.count}`,
            null,
            cardActions
        );
        card.id = context.activity.replyToId;
        const message = MessageFactory.attachment(card);
        message.id = context.activity.replyToId;
        await context.updateActivity(message);
    }

    async sendWelcomeCard(context, cardActions) {
        const initialValue = {
            count: 0
        };
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: initialValue,
            text: 'UpdateCardAction'
        });
        const card = CardFactory.heroCard(
            'Welcome card',
            '',
            null,
            cardActions
        );
        await context.sendActivity(MessageFactory.attachment(card));
    }

    async getSingleMember(context) {
        var member;
        try {
            member = await TeamsInfo.getMember(context, context.activity.from.id);
        }
        catch (e) {
            if(e.code === 'MemberNotFoundInConversation') {
                context.sendActivity(MessageFactory.text('Member not found.'));
                return;
            }
            else {
                console.log(e);
                throw e;
            }
        }
        const message = MessageFactory.text(`You are: ${member.name}.`);
        await context.sendActivity(message);
    }

    async mentionActivityAsync(context) {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }

    async deleteCardActivityAsync(context) {
        await context.deleteActivity(context.activity.replyToId);
    }

    // If you encounter permission-related errors when sending this message, see
    // https://aka.ms/BotTrustServiceUrl
    async messageAllMembersAsync(context) {
        const members = await this.getPagedMembers(context);

        members.forEach(async (teamMember) => {
            const message = MessageFactory.text(`Hello ${ teamMember.givenName } ${ teamMember.surname }. I'm a Teams conversation bot.`);

            var ref = TurnContext.getConversationReference(context.activity);
            ref.user = teamMember;

            await context.adapter.createConversation(ref,
                async (t1) => {
                    const ref2 = TurnContext.getConversationReference(t1.activity);
                    await t1.adapter.continueConversation(ref2, async (t2) => {
                        await t2.sendActivity(message);
                    });
                });
        });

        await context.sendActivity(MessageFactory.text('All messages have been sent.'));
    }

    async getPagedMembers(context) {
        var continuationToken;
        var members = [];
        do {
            var pagedMembers = await TeamsInfo.getPagedMembers(context, 100, continuationToken);
            continuationToken = pagedMembers.continuationToken;
            members.push(...pagedMembers.members);
        } while (continuationToken !== undefined);
        return members;
    }
    
    async connectDb() {
        const { database } = await cosmosClient.databases.createIfNotExists({ id: databaseId });
        const { container } = await database.containers.createIfNotExists({ id: containerId });
    
        return { database: database, container: container };
    }

    async enqueueQuestion(context) {
        const receivedMsg = context.activity;
        let db = null;
        let message = '';

        try{
            db = await this.connectDb();
        } catch (err) {
            console.log(`Error connecting to database: ${err}`);
        }

        try {
            await db.container.items.create({message: receivedMsg, status: "unanswered"});
        } catch (err) {
            console.log(`Error adding new question in db: ${err}`);
            throw err;
        }
    
        message = MessageFactory.text(`Your request has been added to a queue. We will notify you when it is your turn to speak. 😎`);
        // TODO: There are currently x question in the queue. You are y in line.
        await context.sendActivity(message);
    }

    async dequeueQuestion(context) {
        let followupText = 'You have reach the end of the question queue. Yay! 🙌';
        const receivedMsg = context.activity;
        let db = null;
        let message = '';

        try{
            db = await this.connectDb();
        } catch (err) {
            console.log(`Error connecting to database: ${err}`);
        }

        try {
            var { resources: questions } = await db.container.items.query("SELECT * from c WHERE c.status='unanswered'").fetchAll();
        } catch (err) {
            console.log(`Error getting questions from db: ${err}`);
            throw err;
        }

        if (!questions || questions.length == 0) {
            message = MessageFactory.text(followupText);
        } else {
            const currentQuestion = questions.shift();
            // mark question as answered
            try{
                currentQuestion.status = "answered";
                await db.container.item(currentQuestion.id).replace(currentQuestion);
            } catch (err) {
                console.log(`Error marking question as answered in db: ${err}`);
                throw err;
            }

            if (currentQuestion) {
                // replace @mention botname
                const currentQuestionText = currentQuestion.message.text.replace(/<at[^>]*>(.*?)<\/at> *(&nbsp;)*QQ */, '');
                const numberQuestionLeft = questions.length;

                if (numberQuestionLeft > 0) {
                    followupText = 'You have ' + numberQuestionLeft + ' more questions in the queue.';
                }
                message = MessageFactory.text(`🤓From: @${currentQuestion.message.from.name}\n\n🦒Question: ${currentQuestionText}\n\n👀${followupText}`);
                
            } else {
                message = MessageFactory.text(followupText);
            }
        }
        await context.sendActivity(message);
    }
}

module.exports.TeamsConversationBot = TeamsConversationBot;

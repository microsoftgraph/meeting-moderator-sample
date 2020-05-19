import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {BatchRequestStep, BatchRequestContent} from '@microsoft/microsoft-graph-client'
import { getClient } from './graph';

export const getChatIdFromOnlineMeeting = (onlineMeetingUri: string) => {
    const match = onlineMeetingUri.match(/meetup-join\/(.*)\//);
    if (match.length > 1) {
        return match[1];
    }

    return null;
}

export const getChatParticipantsForOnlineMeeting = async (onlineMeetingUri: string) => {
    try {
        let chatId = getChatIdFromOnlineMeeting(onlineMeetingUri);
        const members = await getClient()?.api(`me/chats/${chatId}/members`).version('beta').get();
        if (members && members.value) {
            return members.value;
        }
    } catch (e) {
        return null;
    }

    return null;
}

export const createOnlineMeeting = async (startDateTime: string, endDateTime: string, subject: string) => {
    try {
        return await getClient().api('/me/onlineMeetings').post({
            startDateTime: startDateTime + '-00:00',
            endDateTime: endDateTime + '-00:00',
            subject
        });
    } catch (e) {
        return null;
    }
}

export const createOnlineMeetingsForGroups = async (startDateTime: string, endDateTime: string, numOfGroups: number) => {
    
    let requests: BatchRequestStep[] = [];

    for (let i = 0; i < numOfGroups; i++) {

        const content = {
            startDateTime: startDateTime + '-00:00',
            endDateTime: endDateTime + '-00:00',
            subject: `Group ${i+1} breakout`
        }

        const request = new Request(`/me/onlineMeetings`, {
            method: 'POST',
            body: JSON.stringify(content),
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        requests.push({
            id: i.toString(),
            request
        })
    }

    let responses = [];

    while (requests.length > 0) {

        // batch supports 20 requests at a time
        const currentRequests = requests.splice(0, 20);
        const content = await (new BatchRequestContent(currentRequests)).getContent()

        try {
            const response = await getClient().api('/$batch').post(content);

            for (const r of response.responses) {
                responses[r.id] = r.body;
            }
        } catch (e) {

        }
    }

    return responses;
}

export const sendMessageToOnlineMeeting = async (onlineMeetingUri: string, message: string, mentioned?: MicrosoftGraph.User[]) => {
    try {
        let chatId = getChatIdFromOnlineMeeting(onlineMeetingUri);
        const messageObj: any = {};

        if (mentioned) {
            message += '<br /><br />'
            let mentions = [];
            for (let i = 0; i < mentioned.length; i++) {
                const person = mentioned[i];
                message += ` <at id="${i}">${person.displayName}</at> `;
                mentions.push({
                    "id": i,
                    "mentionText": person.displayName,
                    "mentioned": {
                        "user": {
                            "displayName": person.displayName,
                            "id": person.id,
                            "userIdentityType": "aadUser"
                        }
                    }
                })
            }

            messageObj.mentions = mentions;
        }

        messageObj.body = {
            content: message,
            contentType: 'html'
        }


        await getClient()?.api(`/chats/${chatId}/messages`).version('beta').post(messageObj);
    } catch (e) {
        return null;
    }
}
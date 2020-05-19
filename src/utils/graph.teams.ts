import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {ResponseType, BatchRequestStep, BatchRequestContent} from '@microsoft/microsoft-graph-client'
import { GroupInfo } from './types';
import { getClient } from './graph';

const timeout = (ms: number) => {
    return new Promise(resolve => setTimeout(resolve, ms));
}

export const createTeamAndChannelsFromGroups = async (groups: MicrosoftGraph.User[][], teamName: string) => {
    const team = {
        "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
        "visibility": "private",
        "displayName": teamName,
        "description": "This is a team used for breakout discussions",
        channels: [],
        // installedApps: [
        //     {
        //         // TODO: id here needs to be the one from the Teams store. 
        //         // Upload the manifest/app to the store (for Contoso)
        //         // then find the app in the store and click on the ...
        //         // click copy link - the link will contain the id of the app
        //         'teamsApp@odata.bind':  "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teams.moderator')"
        //     }
        // ]
    };

    for (let i = 0; i < groups.length; i++) {
        // const group = groups[i];
        team.channels.push({
            "displayName": `Group ${i+1}`
        })
    }

    let response: Response;
    try {
        response =  await getClient().api(`teams`)
                                    .version('beta')
                                    .responseType(ResponseType.RAW).post(team);
    } catch (e) {
        return null;
    }
    
    if (!response.ok) {
        return null;
    }                
    
     // need to check the status of our job
    // https://docs.microsoft.com/en-us/graph/api/resources/teamsasyncoperationstatus?view=graph-rest-beta
    const location = response.headers.get('Location')
    let status = 'inProgress';
    let operationResult = null;

    
    while (status === 'inProgress' || status === 'notStarted') {
        // need to wait to ensure team and channels are created
        await timeout(10000);
        try {
            operationResult = await getClient().api(location).version('beta').get();
            status = operationResult.status;
        } catch (e) { }
    }

    if (status !== 'succeeded') {
        // something went wrong creating the team
        return null;
    }            
    
    return operationResult.targetResourceId;
}

export const addUserToTeam = async (teamId: string, userId: string) => {
    try {
        await getClient().api(`groups/${teamId}/members/$ref`).post({
            "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${userId}`
        })
    } catch (e) {
        return null;
    }
}

export const addUsersToTeam = async (teamId: string, userIds: string[]) => {

    let requests :BatchRequestStep[] = [];

    for (const i in userIds) {
        const id = userIds[i];

        const content = {
            "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${id}`
        }

        const request = new Request(`/groups/${teamId}/members/$ref`, {
            method: 'POST',
            body: JSON.stringify(content),
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        requests.push({
            id: i,
            request
        })
    }

    while (requests.length > 0) {

        // batch supports 20 requests at a time
        const currentRequests = requests.splice(0, 20);
        const content = await (new BatchRequestContent(currentRequests)).getContent()

        try {
            await getClient().api('/$batch').post(content);
        } catch (e) {}
    }
}

// export const addUsersToTeam = async (teamId: string, userIds: string[]) => {

//     while (userIds.length > 0) {
//         const currentUsers = userIds.splice(0, 20);

//         let content = {
//             'members@odata.bind': currentUsers.map(id => `https://graph.microsoft.com/v1.0/directoryObjects/${id}`)
//         }

//         try {
//             await getClient().api(`groups/${teamId}`).patch(content);
//         } catch (e) {}
//     }
// }

// export const addUsersToTeam = async (teamId: string, userIds: string[]) => {
//     for (const id of userIds) {
//         await addUserToTeam(teamId, id);
//     }
// }

export const getChannelsForTeam: (teamId: string) => Promise<MicrosoftGraph.Channel[]> = async (teamId: string) => {
    try {
        return (await getClient().api(`teams/${teamId}/channels`).get()).value;
    } catch (e) {
        return null;
    }
}

export const sendMessageToChannel = async (teamId: string, channelId: string, onlineMeetingUrl: string, mentioned: MicrosoftGraph.User[]) => {
    try {
        const response = await getClient()
            .api(`/teams/${teamId}/channels/${channelId}/messages`)
            .version('beta')
            .post(getMessageContentForChannel(onlineMeetingUrl, mentioned));
        return response;
    } catch (e) {
        return null;
    }
}

export const sendMessageToChannels = async (teamId: string, groups: GroupInfo[]) => {
    let requests: BatchRequestStep[] = [];

    for (const i in groups) {
        const group = groups[i]

        const content = getMessageContentForChannel(group.onlineMeeting, group.members);

        const request = new Request(`/teams/${teamId}/channels/${group.id}/messages`, {
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

    while (requests.length > 0) {

        // batch supports 20 requests at a time
        const currentRequests = requests.splice(0, 20);
        const content = await (new BatchRequestContent(currentRequests)).getContent()

        try {
            await getClient().api('/$batch').version('beta').post(content);
        } catch (e) {}
    }
}

export const archiveTeam = async (teamId) => {
    try {
        await getClient()?.api(`/teams/${teamId}/archive`).post({});
    } catch (e) {
        return null;
    }
}

const getMessageContentForChannel = (onlineMeetingUrl: string, mentioned: MicrosoftGraph.User[]) => {
    
    let messageContent = `<h1>Hey everyone!</h1>
Let's use this meeting to have a private breakout!
<br />
<h2><a href="${onlineMeetingUrl}">Join meeting</a></h2>
<br />
`;

    let mentions = [];
    for (let i = 0; i < mentioned.length; i++) {
        const person = mentioned[i];
        messageContent += ` <at id="${i}">${person.displayName}</at> `;
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

    return {
        "body": {
            "contentType": "html",
            "content": messageContent
        },
        "mentions": mentions
    }
}
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// TODO - graph types package is not up to date with the API
export type GraphEvent = MicrosoftGraph.Event & {
    isOnlineMeeting: boolean,
    onlineMeeting: {
        joinUrl: string
    },
    onlineMeetingProvider: 'teamsForBusiness' | 'skypeForBusiness' | 'skypeForConsumer'
}

export type Team = microsoftgraphbeta.Team & {
    'template@odata.bind': string
}

export interface GroupInfo {
    id: string,
    name: string,
    onlineMeeting: string,
    members: MicrosoftGraph.User[]
}

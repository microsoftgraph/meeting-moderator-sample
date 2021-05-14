import { TeamsHelper } from "@microsoft/mgt-element"
import * as MicrosoftTeams from '@microsoft/teams-js';

export const openTeamsUrl = (url: string) => {
    if (TeamsHelper.isAvailable) {
        MicrosoftTeams.initialize(() => {
            console.log(url)
            MicrosoftTeams.executeDeepLink(url, (success) => {
                if (!success) {
                    window.open(url, '_blank');
                }
            });
        })
    } else {
        window.open(url, '_blank')
    }
}
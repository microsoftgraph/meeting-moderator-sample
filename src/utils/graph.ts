import { Providers, ProviderState } from '@microsoft/mgt-element';

export const getClient = () => {
    if (Providers.globalProvider.state === ProviderState.SignedIn) {
        return Providers.globalProvider.graph.client;
    }

    return null;
}

export const getCurrentSignedInUser = async () => {
    try {
        return await getClient()?.api('me').get();
    } catch (e) {
        return null;
    }
}
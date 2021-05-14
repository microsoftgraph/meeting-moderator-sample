import React from 'react';
import { TeamsProvider } from '@microsoft/mgt-teams-provider';

export function TeamsAuth() {
    TeamsProvider.handleAuth();

    return (
        <div>Now signing you in!</div>
    );
}
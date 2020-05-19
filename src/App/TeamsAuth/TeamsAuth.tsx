import React from 'react';
import { TeamsProvider } from '@microsoft/mgt';

export function TeamsAuth() {
    TeamsProvider.handleAuth();

    return (
        <div>Now signing you in!</div>
    );
}
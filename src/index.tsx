import React from 'react';
import ReactDOM from 'react-dom';
import App from './App/App';

import {Msal2Provider} from '@microsoft/mgt-msal2-provider';
import {TeamsProvider} from '@microsoft/mgt-teams-provider';
import { Providers } from '@microsoft/mgt-element';

import * as MicrosoftTeams from "@microsoft/teams-js";

TeamsProvider.microsoftTeamsLib = MicrosoftTeams;

let provider;
const clientId = 'eb3b6c7d-41fa-4607-b010-3ddd0a1f071b';

const scopes = [ 
  'user.read',
  'people.read',
  'user.readbasic.all',
  'contacts.read',
  'calendars.read',
  'Presence.Read.All',
  'Presence.Read'
]

if (TeamsProvider.isAvailable) {
  provider = new TeamsProvider({
    clientId,
    scopes,
    authPopupUrl: '/teamsauth'
  })
} else {
  provider = new Msal2Provider({
    clientId,
    scopes,
    redirectUri: window.location.origin
  });
}

Providers.globalProvider = provider;


ReactDOM.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
  document.getElementById('root')
);

// serviceWorker.register();
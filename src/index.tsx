import React from 'react';
import ReactDOM from 'react-dom';
import App from './App/App';
import * as serviceWorker from './serviceWorker';

import {Providers, MsalProvider, TeamsProvider} from '@microsoft/mgt';
import * as MicrosoftTeams from "@microsoft/teams-js";

TeamsProvider.microsoftTeamsLib = MicrosoftTeams;

let provider;
const clientId = '172ac0f4-104f-4765-9ccf-3df699905899';

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
  provider = new MsalProvider({
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

serviceWorker.register();
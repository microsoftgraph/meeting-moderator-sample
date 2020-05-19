import React, { useState } from 'react';
import './App.css';

import { Providers, ProviderState, TeamsHelper } from '@microsoft/mgt';
import { Switch, Route, BrowserRouter, useHistory } from "react-router-dom";
import { MgtTemplateProps, Login, Person } from '@microsoft/mgt-react';
import { initializeIcons, IconButton, Spinner, SpinnerSize } from '@fluentui/react';
import { wrapWc } from 'wc-react';
import '@pwabuilder/pwainstall';

import { AgendaView } from './AgendaView/AgendaView';
import { EventView } from './EventView/EventView';
import { TeamsAuth } from './TeamsAuth/TeamsAuth';
import { TabConfig } from './TeamsTabConfig/TabConfig';
import {ReactComponent as BicycleImage} from '../images/tandem-bicycle.svg';

function App() {

  initializeIcons();

  const [authState, setAuthState] = useState(Providers.globalProvider.state);
  const PwaInstall = wrapWc('pwa-install');

  if (window.location.pathname.startsWith('/teamsauth')) {
    return <TeamsAuth />
  }

  Providers.onProviderUpdated(() => {
    setAuthState(Providers.globalProvider.state)
  })

  return (
    <BrowserRouter>

    <div className="App">
      {!TeamsHelper.isAvailable && 
        <header className="App-header">
          <HeaderTitle />
          <div className="InstallButton">
            <PwaInstall></PwaInstall>
          </div>
            <div className="HelpIcon">
              <IconButton 
                className="AboutButton"
                iconProps={{iconName: 'Help'}}
                onClick={(e) => window.open("https://github.com/microsoftgraph/meeting-moderator-sample", '_blank')}></IconButton>
          </div>
          {authState !== ProviderState.SignedOut && 
            <Login>
              <SignedInButtonContent template="signed-in-button-content" />
            </Login>
          }
        </header>
      }
      <div className="App-content">
        {authState === ProviderState.SignedIn && <MainContent />}
        {authState === ProviderState.SignedOut && <LoginPage />}
        {authState === ProviderState.Loading && <Spinner size={SpinnerSize.large} />}
      </div>
    </div>
    </BrowserRouter>
  );
}

const MainContent = () => {
  return (
    <Switch>
      <Route path="/events/:id">
        <EventView />
      </Route>
      <Route path="/teamsconfig">
        <TabConfig />
      </Route>
      <Route path="/">
        <AgendaView />
      </Route>
    </Switch>
  )
}

const SignedInButtonContent = (props: MgtTemplateProps) => {
  const {personDetails, personImage} = props.dataContext;

  return <div className="SignedInButtonContent" >
    <Person personDetails={personDetails} personImage={personImage}></Person>
  </div>
}

const HeaderTitle = () => {

  const history = useHistory();

  const handleTitleClick = () => {
    if (history.location.pathname !== '/') {
      history.push(`/`);
    }
  }

  return <div className="Event-title">
    <div className="App-title" onClick={handleTitleClick} >Moderator</div>
  </div>;
}

const LoginPage = () => {
  return <div className="Card welcome">
    <BicycleImage />
    <h1>Moderator</h1>
    <div className="app-description">
      Select an online meeting to start moderating. As moderator, you can split your meeting participants into breakout rooms.
    </div>
    <Login />
  </div>
}

export default App;

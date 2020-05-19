import React from 'react';

import { useHistory } from "react-router-dom";
import { ActionButton } from '@fluentui/react';
import { PersonCardInteraction, Providers } from '@microsoft/mgt';
import { Agenda, MgtTemplateProps, People } from '@microsoft/mgt-react';

import { getDateHeader, getFormatedTime } from '../../utils/date';
import { GraphEvent } from '../../utils/types';
import './AgendaView.css';
import { openTeamsUrl } from '../../utils/helpers';

export function AgendaView() {

  const history = useHistory();
    
  const handleEventClick = async (event: GraphEvent) => {

    let token = await Providers.globalProvider.getAccessTokenForScopes(
      'Calendars.ReadWrite',
      'Chat.ReadWrite',
      'Group.ReadWrite.All',
      'OnlineMeetings.ReadWrite',
      'GroupMember.ReadWrite.All',);

      if (token) {
        history.push(`/events/${event.id}`);
      }
  }

  return (
    <div className="AgendaView">

      <div className="HelloView">
        <div className="Title">Hello!</div>
        <div>Select a meeting to start moderating!</div>
      </div>

      <div className="Card">
        <Agenda days={5} groupByDay>
          <AgendaHeader template="header"></AgendaHeader>
          <AgendaEvent template="event" onClick={handleEventClick}></AgendaEvent>
        </Agenda>
      </div>
    </div>
  );
}

let AgendaHeader = (props: MgtTemplateProps) => {
  let date = new Date(props.dataContext.header);
  return <div className="AgendaHeader">{getDateHeader(date)}</div>
}

let AgendaEvent = (props: MgtTemplateProps & {onClick: (event: GraphEvent) => void}) => {

  const event: GraphEvent = props.dataContext.event;
  const start = new Date(event.start.dateTime)
  const end = new Date(event.end.dateTime);

  const handleJoin = () => {
    openTeamsUrl(event.onlineMeeting.joinUrl);
}

  return (
    <div className={`AgendaEvent ${event.isOnlineMeeting ? '' : 'OfflineEvent'}`}>
      <div className="AgendaEventDetails">
        <div className="EventTime">{`${getFormatedTime(start)} - ${getFormatedTime(end)}`}</div>
        <div className="EventSubject">{event.subject}</div>
        <div className="EventAttendees">
          <People 
            peopleQueries={event.attendees.map(a => a.emailAddress.address)} 
            personCardInteraction={event.isOnlineMeeting ? PersonCardInteraction.hover : PersonCardInteraction.none}
            showPresence={event.isOnlineMeeting}>
          </People>
        </div>
      </div>
      {event.isOnlineMeeting && 
        <div className="AgendaActions">
          <ActionButton 
            iconProps={{ iconName: 'Group', style: {fontSize: 24} }} 
            text="Moderate" 
            style={{fontWeight: 600}}
            onClick={() => props.onClick(event)}/>
          <ActionButton 
            iconProps={{ iconName: 'TeamsLogo16', style: {fontSize: 24} }}
            text="Join"
            style={{fontWeight: 600}}
            onClick={handleJoin}/>
        </div>
      }
    </div>
  )
}
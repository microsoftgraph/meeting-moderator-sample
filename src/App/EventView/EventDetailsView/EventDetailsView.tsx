import React from 'react';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Person, PersonCardInteraction, PersonViewType } from '@microsoft/mgt-react';
import { ActionButton } from '@fluentui/react';

import { GraphEvent } from '../../../utils/types';
import './EventDetailsView.css'
import { openTeamsUrl } from '../../../utils/helpers';

export const EventDetailsView = (props: {event: GraphEvent, participants: MicrosoftGraph.User[]}) => {

    const handleJoin = () => {
        openTeamsUrl(props.event.onlineMeeting.joinUrl);
    }

    return <div>
        <div className="EventDetailsContainer Card">
            <div className="CardTitle">Details</div>
            <div className="EventDetails">
                <div className="EventDetail">
                    <div className="EventKey">Subject</div>
                    <div className="EventValue">{props.event.subject}</div>
                </div>
                <div className="EventDetail">
                    <div className="EventKey">Date</div>
                    <div className="EventValue">TODO</div>
                </div>
                <div className="EventDetail">
                    <div className="EventKey">Time</div>
                    <div className="EventValue">TODO</div>
                </div>
                <div className="EventDetail">
                    <div className="EventKey">Body</div>
                    <div className="EventValue">{props.event.bodyPreview}</div>
                </div>
                <div className="EventDetail JoinButton">
                    <div className="EventKey"></div>
                    <div className="EventValue">
                        <ActionButton iconProps={{ iconName: 'TeamsLogo16' }} text="Join Meeting" onClick={handleJoin}/>
                    </div>
                </div>
            </div>

        </div>
        <div className="ParticipantsContainer Card">
            <div className="CardTitle">Participants</div>
            <div className="ParticipantsList">
                {props.participants.map((p,i) => (
                    <div className="Participant" key={i}>
                        <Person personDetails={p} showPresence avatarSize="large" fetchImage view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.hover} />
                    </div>
                ))}
            </div>
        </div>
    </div>
}
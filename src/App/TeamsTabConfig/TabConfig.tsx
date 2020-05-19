import React, { useState, useEffect } from 'react';

import * as MicrosoftTeams from "@microsoft/teams-js";
import { Icon } from '@fluentui/react';

import { GraphEvent } from '../../utils/types';
import { getEvents } from '../../utils/graph.events';

export const TabConfig = () =>  {
    const [message, setMessage] = useState('Looking for event...');

    useEffect(() => {
        MicrosoftTeams.getContext(async (context) => {
            const chatId = context.chatId;
            const events : GraphEvent[] = await getEvents();

            if (events) {
                let filteredEvents = events.filter(e => e.onlineMeeting && unescape(e.onlineMeeting.joinUrl).includes(chatId));
                
                if (filteredEvents.length > 0) {
                    setMessage(`Found event ${filteredEvents[0].subject}. Click "Save" to add tab!`)
                    MicrosoftTeams.settings.setSettings({"entityId": 'moderator', "contentUrl": `${window.location.origin}/events/${filteredEvents[0].id}`});
                    MicrosoftTeams.settings.setValidityState(true);
                } else {
                    setMessage('Event not found :(. This chat might not be a meeting chat.')
                }
            } else {
                setMessage('No events found in calendar or there was error getting events...')
            }
        });
    }, [])

    return (
        <div>
            <Icon iconName="ConfigurationSolid" className="Logo" />
            <h1>Setting up the Moderator tab</h1>
            {message}
        </div>
    );
}
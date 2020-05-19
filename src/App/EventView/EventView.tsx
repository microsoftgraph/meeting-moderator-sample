import React, { useState, useEffect } from 'react';

import { Pivot, PivotItem, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { BreakoutView } from './BreakoutView/BreakoutView';
import { useParams } from 'react-router-dom';
import { QuestionQueueView } from './QuestionQueueView/QuestionQueue';
import { EventDetailsView } from './EventDetailsView/EventDetailsView';
import { getEventFromId } from '../../utils/graph.events';
import { getChatParticipantsForOnlineMeeting } from '../../utils/graph.onlineMeetings';
import { GraphEvent } from '../../utils/types';



export function EventView() {
    let { id } = useParams();

    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);

    const [event, setEvent] = useState<GraphEvent>(null);
    const [participants, setParticipants] = useState(null);

    useEffect(() => {
        (async () => {
            if (isLoading) {

                const event: GraphEvent = await getEventFromId(id);

                if (!event) {
                    setError('Event not found!');
                    return;
                }

                if (event && (!event.isOnlineMeeting || event.onlineMeetingProvider !== 'teamsForBusiness')) {
                    setError('Event is not a Teams meeting');
                    return;
                }

                let attendees = await getChatParticipantsForOnlineMeeting(event.onlineMeeting.joinUrl);
                
                setEvent(event);
                setParticipants(attendees);
                setIsLoading(false);
            }
        })();
    }, []);

    if (error) {
        return <MessageBar messageBarType={MessageBarType.severeWarning}>{error}</MessageBar>;
    }

    if (isLoading) {
        return (
        <div>
            <Spinner size={SpinnerSize.large} 
                        labelPosition="bottom"/>
        </div>)
        ;
    }

    return (
        <Pivot defaultSelectedKey='0'>
            <PivotItem headerText='Breakouts'>
                <BreakoutView event={event} attendees={participants}></BreakoutView>
            </PivotItem>
            <PivotItem headerText='Question Queue'>
                <QuestionQueueView event={event} />
            </PivotItem>
            <PivotItem headerText='Details'>
                <EventDetailsView event={event} participants={participants}></EventDetailsView>
            </PivotItem>
        </Pivot>
    );
}


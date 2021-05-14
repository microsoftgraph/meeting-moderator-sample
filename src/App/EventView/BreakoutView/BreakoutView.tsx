import React, { useState, useEffect } from 'react';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as FluentUI from '@fluentui/react';
import { Person, PersonCardInteraction, PersonViewType } from '@microsoft/mgt-react';

import { GraphEvent, GroupInfo } from '../../../utils/types';
import { getEventExtension, updateEventExtension } from '../../../utils/graph.events';
import { getCurrentSignedInUser } from '../../../utils/graph';
import { createOnlineMeetingsForGroups, sendMessageToOnlineMeeting, getChatIdFromOnlineMeeting } from '../../../utils/graph.onlineMeetings';
import { createTeamAndChannelsFromGroups, addUsersToTeam, getChannelsForTeam, sendMessageToChannels, archiveTeam } from '../../../utils/graph.teams';

import { BreakoutsCreatorView } from './BreakoutsCreatorView/BreakoutsCreatorView';
import './BreakoutView.css';
import { openTeamsUrl } from '../../../utils/helpers';

export interface BreakoutsInfo {
    teamName: string,
    teamId: string,
    moderators: MicrosoftGraph.User[],
    groups: GroupInfo[]
}

export interface BreakoutViewProps {
    event: GraphEvent,
    attendees: MicrosoftGraph.User[]
}

export function BreakoutView(props: BreakoutViewProps) {

    const [isLoading, setIsLoading] = useState(true);
    const [loadingMessage, setLoadingMessage] = useState(null);


    const [moderators, setModerators] = useState(null);
    const [participants, setParticipants] = useState(null);
    const [breakouts, setBreakouts] = useState<BreakoutsInfo>(null);

    const [currentSignedInUser, setCurrentSignedInUser] = useState(null);

    useEffect(() => {
        (async () => {
            setIsLoading(true);
            
            const currentSignedInUser = await getCurrentSignedInUser();
            setCurrentSignedInUser(currentSignedInUser);

            let moderators = [currentSignedInUser];

            const extension = await getEventExtension(props.event.id);
            if (extension && extension.breakouts && extension.breakouts !== '') {
                const breakouts : BreakoutsInfo = JSON.parse(extension.breakouts);
                moderators = breakouts.moderators;
                
                setBreakouts(breakouts);
            } 

            const participants = props.attendees.filter(p => !moderators.find(m => m.id === p.id))

            setParticipants(participants);
            setModerators(moderators);
            setIsLoading(false);
        })();
    }, []);

    if (isLoading) {
        return (
        <div>
            <FluentUI.Spinner 
                size={FluentUI.SpinnerSize.large} 
                label={loadingMessage} 
                labelPosition="bottom"/>
        </div>)
        ;
    }

    const handleModeratorsAdded = (person) => {
        setModerators([...moderators, person]);
        setParticipants(participants.filter(p => p.id !== person.id))
    }

    const handleModeratorsRemoved = (person) => {
        setModerators(moderators.filter(m => m.id !== person.id));
        setParticipants([person, ...participants]);
    }

    const handleCreateGroups = async (groups: MicrosoftGraph.User[][]) => {
        setIsLoading(true);

        const teamName = props.event.subject + ' Breakouts';

        setLoadingMessage(`Creating Team "${teamName}"` )
        const teamId = await createTeamAndChannelsFromGroups(groups, teamName);
        if (!teamId) return;

        const userIds = [...participants, ...moderators]
            .filter(u => u.id !== currentSignedInUser.id)
            .map(u => u.id);
        setLoadingMessage(`Adding users to team`)
        await addUsersToTeam(teamId, userIds);

        const channels = await getChannelsForTeam(teamId);
        if (!channels) return;
        
        setLoadingMessage(`Creating online meetings`)
        const onlineMeetings = await createOnlineMeetingsForGroups(
            props.event.start.dateTime, 
            props.event.end.dateTime,
            groups.length);


        setLoadingMessage(`Sending messages to each channel`)
        const groupInfo: GroupInfo[] = [];

        for (const channel of channels) {

            const match = channel.displayName.match(/Group ([0-9]+)/);
            if (match && match.length > 1) {
                const index = Number.parseInt(match[1]) - 1;
                const groupName = `Group ${index + 1}`;

                const group = groups[index];
                const onlineMeeting = onlineMeetings[index];

                groupInfo.push({
                    name: groupName,
                    id: channel.id,
                    onlineMeeting: onlineMeeting.joinUrl,
                    members: group.map(m => {return {id: m.id, displayName: m.displayName}})
                });
            }   
        }

        await sendMessageToChannels(teamId, groupInfo);

        const breakouts: BreakoutsInfo = {
            teamName,
            teamId,
            moderators: moderators.map(m => {return {id: m.id, displayName: m.displayName}}),
            groups: groupInfo
        };

        updateEventExtension(props.event.id, breakouts);
        setBreakouts(breakouts);
        setIsLoading(false);
    }
    
    const archiveBreakouts = async () => {
        setLoadingMessage('Archiving team');
        setIsLoading(true);
        await archiveTeam(breakouts.teamId);
        await updateEventExtension(props.event.id, null)
        setBreakouts(null);
        setIsLoading(false);
    };


    return (
    <div className="Breakouts">
        {!breakouts ?
            <BreakoutsCreatorView 
                participants={participants} 
                moderators={moderators}
                onModeratorsAdded={handleModeratorsAdded}
                onModeratorsRemoved={handleModeratorsRemoved}
                currentSignedInUser={currentSignedInUser}
                onCreateClick={handleCreateGroups}/>
        : 
            <Breakouts 
                breakouts={breakouts} 
                event={props.event} 
                onArchive={archiveBreakouts}></Breakouts>
        }
    </div>
    );
}

const Breakouts = (props: {breakouts: BreakoutsInfo, event: GraphEvent, onArchive: () => void}) => {

    const [message, setMessage] = useState('Hey [group-name], you have five minutes left! [meeting-link]')
    const [isLoading, setIsLoading] = useState(false);
    const [showDialog, setShowDialog] = useState(false);

    const sendMessage = async () => {
        setIsLoading(true);

        let msg = message.replace('[meeting-link]', `<a href="${props.event.onlineMeeting.joinUrl}">Join main meeting</a>`)
        for (const {onlineMeeting, members, name} of props.breakouts.groups) {
            let groupMsg = msg.replace('[group-name]', name)
            await sendMessageToOnlineMeeting(onlineMeeting, groupMsg, members);
        }
        setIsLoading(false);
    }

    return (
        <div className="Breakouts Card">
            <div className="BreakoutsTitleContainer">
                <div className="CardTitle">Send message to all breakouts</div>
                <div className="BreakoutsArchiveButton">
                    <FluentUI.DefaultButton text="Archive" onClick={() => setShowDialog(true)} style={{color:'red', borderColor:'red'}} />
                </div>
            </div>
            <div className="MessageTextField">
                <FluentUI.TextField 
                    label="Send a message to all breakouts" 
                    multiline rows={3} 
                    value={message}
                    onChange={(e, value) => setMessage(value)} />
            </div>
            <div className="PlacedholderDescription">
                <FluentUI.Label>The following placeholders will be replaced in the sent message</FluentUI.Label>
                <div> - [meeting-link]: link to original meeting</div>
                <div> - [group-name]: name of the group where the message is sent</div>
            </div>
            <div className="MessageSendButtonContainer">
                {isLoading
                    ? <FluentUI.Spinner className="MessageSpinner" size={FluentUI.SpinnerSize.small} label="Sending Message" labelPosition="right"/>
                    : <FluentUI.PrimaryButton  text="Send" disabled={!message} onClick={sendMessage}></FluentUI.PrimaryButton>
                }
            </div>
            <div className="BreakoutGroups">
                {props.breakouts.groups.map((group, i) => 
                    <div key={i}>
                        <BreakoutGroup group={group}></BreakoutGroup>
                    </div>
                )}
            </div>

            

            <FluentUI.Dialog
                hidden={!showDialog}
                onDismiss={() => setShowDialog(false)}
                dialogContentProps={{
                    type: FluentUI.DialogType.normal,
                    title: 'Archive Breakout Groups!',
                    closeButtonAriaLabel: 'Archive',
                    subText: 'Are you sure you want to archive the breakout groups?',
                }}
                modalProps={{
                    isBlocking: false,
                    styles: { main: { maxWidth: 450 } }
                }}>
                <FluentUI.DialogFooter>
                    <FluentUI.PrimaryButton onClick={props.onArchive} text="Archive" style={{background: 'red', border: 'red'}} />
                    <FluentUI.DefaultButton onClick={() => setShowDialog(false)} text="NOOO, don't Archive!" />
                </FluentUI.DialogFooter>
            </FluentUI.Dialog>
        </div>
    );
}

const BreakoutGroup = (props: {group: GroupInfo}) => {

    const handleJoin = () => {
        openTeamsUrl(props.group.onlineMeeting);
    }

    const handleChat = () => {
        let chatId = getChatIdFromOnlineMeeting(props.group.onlineMeeting);
        let url = `https://teams.microsoft.com/l/chat/${chatId}`;
        openTeamsUrl(url);
    }

    return (
        <div className="GroupView">
            <div className="GroupTitle">{props.group.name}</div>
            <div className="GroupPeople">
                {props.group.members.map((p, i) => 
                    <div className="GroupPerson" key={i}>
                        <Person personDetails={p} showPresence avatarSize="large" fetchImage view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.hover} />
                    </div>
                )}
                <div className="GroupActions">
                    <FluentUI.ActionButton iconProps={{ iconName: 'TeamsLogo16' }} text="Join" onClick={handleJoin}/>
                    <FluentUI.ActionButton iconProps={{ iconName: 'CannedChat' }} text="Chat" onClick={handleChat}/>
                </div>
            </div>
        </div>
        );
}
import React, { useState, useEffect } from 'react';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { PrimaryButton, IconButton, Slider } from '@fluentui/react';
import { Person, PeoplePicker } from '@microsoft/mgt-react';
import { PersonCardInteraction, PersonViewType } from "@microsoft/mgt";

import {ReactComponent as MoveIcon} from '../../../../images/move.svg';
import './BreakoutsCreatorView.css'

export type BreakoutCreatorViewProps = {
    participants: MicrosoftGraph.User[], 
    moderators: MicrosoftGraph.User[],
    onModeratorsAdded: Function,
    onModeratorsRemoved: Function,
    currentSignedInUser: MicrosoftGraph.User,
    onCreateClick: (groups: MicrosoftGraph.User[][]) => void,
}

export const BreakoutsCreatorView = (props: BreakoutCreatorViewProps) => {

    const [groupSize, setGroupSize] = useState(5);
    const [groups, setGroups] = useState<MicrosoftGraph.User[][]>(null);

    useEffect(() => {
        setGroups(generateGroups(props.participants, groupSize));
    }, [groupSize, props.participants]);

    const handleSelectionChanged = (e) => {
        let person = e.detail[0];
        e.target.selectedPeople = [];
        props.onModeratorsAdded(person);
    };

    const handleRemoveClicked = (person) => {
        props.onModeratorsRemoved(person)
    }

    return <div className="BreakoutsCreatorView">
        <div className="Breakouts Card">
            <div className="CardTitle">Create Breakouts</div>
            <Slider
                className="Slider"
                label="Group size"
                min={2}
                max={10}
                step={1}
                defaultValue={groupSize}
                showValue={true}
                onChange={(value: number) => setGroupSize(value)}
                snapToStep
                />
            <div className="Groups">
                {groups && groups.map((g, i) => (<GroupsCreatorGroupView groupMembers={g} name={`Group ${i + 1}`} key={i} />))}
            </div>
            <div className="Actions">
                <PrimaryButton text='Create Breakouts' onClick={() => props.onCreateClick(groups)} />
            </div> 

        </div>
        <div className="Moderators Card">
            <div className="CardTitle">Moderators</div>
            <div className="ModeratorsList">
                {props.moderators.map((p, i) => 
                    <div className="ModeratorPerson Card" key={i}>
                        <Person personDetails={p} showPresence avatarSize="large" fetchImage view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.click} />
                        {props.currentSignedInUser.id === p.id ? '' : (
                            <div className="ModeratorCloseIconButton">
                                <IconButton  iconProps={{iconName: 'ChromeClose'}} onClick={(e) => handleRemoveClicked(p)}></IconButton>
                            </div>
                        )}
                    </div>
                )}
            </div>
            <PeoplePicker people={props.participants} selectionChanged={handleSelectionChanged}></PeoplePicker>
        </div>
    </div>

}

const GroupsCreatorGroupView = (props: {groupMembers: MicrosoftGraph.User[], name: string}) => {

    return (
    <div className="GroupView">
        <div className="GroupTitle">{props.name}</div>
        <div className="GroupPeople">
            {props.groupMembers.map((p, i) => 
                <div className="GroupPerson" key={i}>
                    <MoveIcon className="GroupPersonMoveIcon"></MoveIcon>
                    <Person personDetails={p} showPresence avatarSize="large" fetchImage view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.click} />
                </div>
            )}
        </div>
    </div>
    );
}

/**
 * Shuffles array in place.
 * from: https://stackoverflow.com/questions/6274339/how-can-i-shuffle-an-array
 * @param {Array} a items An array containing the items.
 */
function shuffle(a) {
    var j, x, i;
    for (i = a.length - 1; i > 0; i--) {
        j = Math.floor(Math.random() * (i + 1));
        x = a[i];
        a[i] = a[j];
        a[j] = x;
    }
    return a;
}

const generateGroups = (users: MicrosoftGraph.User[], groupSize: number) => {
    const shuffled = shuffle(users);
    const numberOfGroups = Math.ceil(shuffled.length / groupSize);
    const groups = [];

    for (let i = 0; i < numberOfGroups; i++) {
        groups.push([]);
    }

    let personCount = 0;
    let groupNumber = 0;

    for (const person of shuffled) {
        if (personCount === groupSize) {
            groupNumber++;
            personCount = 0;
        }

        groups[groupNumber].push(person);

        personCount++;
    }

    return groups;
}

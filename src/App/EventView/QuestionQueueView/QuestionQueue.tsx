import React, { useState, useEffect } from 'react';

import { MessageBar, Spinner, MessageBarType, SpinnerSize, PrimaryButton } from '@fluentui/react';
import { Person } from '@microsoft/mgt-react';
import { CosmosClient } from '@azure/cosmos';
import * as signalR from '@microsoft/signalr';
import axios from 'axios';

import { GraphEvent } from '../../../utils/types';
import { getEventExtension } from '../../../utils/graph.events';
import './QuestionQueue.css'

// Cosmos db config
const endpoint = process.env.COSMOS_ENDPOINT || process.env.REACT_APP_COSMOS_ENDPOINT;
const key = process.env.COSMOS_KEY || process.env.REACT_APP_COSMOS_KEY;
const databaseId = process.env.COSMOS_DATABASEID || process.env.REACT_APP_COSMOS_DATABASEID;
const containerId = process.env.COSMOS_CONTAINERID || process.env.REACT_APP_COSMOS_CONTAINERID;
const client = new CosmosClient({ endpoint, key });

// Functions config
const apiBaseUrl = process.env.FUNCTIONS_BASEURL || process.env.REACT_APP_FUNCTIONS_BASEURL;
const axiosConfig = {};

// Get data from cosmos db directly instead of using signal r
async function connectCosmos() {
    const { database } = await client.databases.createIfNotExists({ id: databaseId });
    const { container } = await database.containers.createIfNotExists({ id: containerId });

    return { database, container };
}

async function updateRecordCosmos(item, container) {
    const { id } = item;
    item.status = 'answered';
    await container.item(id).replace(item);
}

interface Question {
    message: any,
    status: any
}


export const QuestionQueueView = (props: {event: GraphEvent}) => {
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);
    const [questions, setQuestionList] = useState<Question[]>(null);

    useEffect(() => {
        (async () => {
            setIsLoading(true);
            const questions = await getMessages();

            const info = await getConnectionInfo();
            let accessToken = info.accessToken;
            const options = {
                accessTokenFactory: () => {
                    if (accessToken) {
                        const _accessToken = accessToken;
                        accessToken = null;
                        return _accessToken;
                    } else {
                        return info.accessToken;
                    }
                }
            }
            const connection = new signalR.HubConnectionBuilder()
                .withUrl(info.url, options)
                .configureLogging(signalR.LogLevel.Information)  
                .build();
        
            connection.on('messageUpdated', (updatedMessage) => {
                const updatedQuestions = messageUpdated(updatedMessage, questions);
                setIsLoading(true);
                setQuestionList(updatedQuestions);
                setIsLoading(false);

            });

            connection.onclose(() => {
                console.log('disconnected');
                setTimeout(() => {
                    startConnection(connection)
                }, 2000);
            });

            startConnection(connection);
            setQuestionList(questions);

            // TODO: Use teamid to reach proper container by id from cosmos db
            const extension = await getEventExtension(props.event.id);
            if (extension && extension.breakouts && extension.breakouts !== '') {
                const teamId = JSON.parse(extension.breakouts).teamId;
            } 
            
            setIsLoading(false);
        })();
    },[]);
    
    if (error) {
        return <MessageBar messageBarType={MessageBarType.severeWarning}>{error}</MessageBar>;
    }

    if (isLoading) {
        return (
        <div>
            <Spinner size={SpinnerSize.large} labelPosition="bottom"/>
        </div>)
        ;
    }

    return <div className="card">
        <h2 className="title">Manage Questions</h2>
        <div className="description">A Bot has been added to your Teams team. It is working hard to get you the questions as shown below.</div>
        <Questions questions={questions}></Questions>
    </div>
}

const Questions = (props: { questions: Question[] }) => {
    // todo: add user id in the json object to prompt Person element
    const disabled = false;

    const handleUpdateStatus = async (question) => {
        const db = await connectCosmos();
        await updateRecordCosmos(question, db.container);
    }

    return(
        <div className="question-list card">
        {props.questions.filter(q => q.status === "unanswered").map(( q, i ) => (            
            <div className="question" key={i}>
                <Person userId={q.message.from.aadObjectId} fetchImage showName avatarSize="large" showPresence></Person>
                <span className="question-text">{q.message.text.replace(/<at[^>]*>(.*?)<\/at> *(&nbsp;)*QQ */, '')}</span>
                <PrimaryButton className="question-answer" text="Answer" onClick={()=>handleUpdateStatus(q)} allowDisabledFocus disabled={disabled} ></PrimaryButton>
            </div>
        ))}
        </div> 
    )
}

function startConnection(connection) {
    console.log('connecting...');
    connection.start()
        .then(() => {console.log('connected!');})
        .catch((err) => {
            console.log('error connecting to signal r: ' + err);
            setTimeout(() => { startConnection(connection)}, 2000);
        })
}

function getMessages() {
    return axios.post(`${apiBaseUrl}/api/getMessages`, null, axiosConfig)
        .then(response => { return response.data})
        .catch((err) => { console.log('error getting message: ' + err); return {}})
}

function getConnectionInfo() {
    return axios.post(`${apiBaseUrl}/api/negotiate`, null, axiosConfig)
    .then(response => { return response.data})
    .catch((err) => { console.log('error getting connection info: ' + err); return {}})
}

function messageUpdated(updatedMessage, questions) {
    const message = questions.find( q => q.id === updatedMessage.id)
    if (message) {
        const index = questions.findIndex( q => q.id === updatedMessage.id);
        questions[index] = updatedMessage;
    } else {
        questions.push(updatedMessage)
    }

    return questions;
}
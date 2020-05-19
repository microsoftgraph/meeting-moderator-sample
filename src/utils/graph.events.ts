import { getClient } from "./graph";

const extensionName = 'com.moderatorTest'

export const getEventFromId = async (id: string) => {
    try {
        return await getClient()
            .api(`me/calendar/events/${id}`)
            .select([
                'id', 
                'subject', 
                'bodyPreview', 
                'onlineMeeting', 
                'end', 
                'start', 
                'isOnlineMeeting', 
                'onlineMeetingProvider'])
            .expand(`extensions($filter=id eq '${extensionName}')`)
            .get();
    } catch (e) {
        return null;
    }
}

export const getEventExtension = async (eventId: string) => {
    try {
        let extension = await getClient().api(`me/events/${eventId}/extensions/${extensionName}`).get();
        return extension;
    } catch (e) {
        if (e.code === 'ErrorItemNotFound') {
            return await createEventExtension(eventId);
        }
        return null;
    }
}

export const createEventExtension = async (eventId: string) => {
    try {
        let extension = await getClient().api(`me/events/${eventId}/extensions`).post({
            "@odata.type": "microsoft.graph.openTypeExtension",
            "extensionName": extensionName,
            "breakouts": ""
        });
        return extension;
    } catch (e) {
        return null;
    }
}

export const updateEventExtension = async (eventId: string, breakouts: any) => {
    const content = {
        "@odata.type": "microsoft.graph.openTypeExtension",
        "breakouts": !!breakouts ? JSON.stringify(breakouts) : ''
    };

    try {
        let extension = await getClient().api(`me/events/${eventId}/extensions/${extensionName}`).patch(content);
        return extension;
    } catch (e) {
        return null;
    }
}

export const getEvents = async (days = 3) => {
    const startDate = new Date();
    startDate.setHours(0, 0, 0, 0);

    const endDate = new Date(startDate.getTime());
    endDate.setDate(startDate.getDate() + days); // next 3 days

    const sdt = `startdatetime=${startDate.toISOString()}`;
    const edt = `enddatetime=${endDate.toISOString()}`;

    try {
        let response =  await getClient()?.api(`me/calendarview?${sdt}&${edt}`).get();
        return response.value;
    } catch (e) {
        return null;
    }
}

export const isToday = (someDate) => {
    const today = new Date()
    return someDate.getDate() === today.getDate() &&
        someDate.getMonth() === today.getMonth() &&
        someDate.getFullYear() === today.getFullYear()
}

export const isTomorrow = (someDate) => {
    const today = new Date()
    today.setDate(today.getDate() + 1);
    return someDate.getDate() === today.getDate() &&
        someDate.getMonth() === today.getMonth() &&
        someDate.getFullYear() === today.getFullYear()
}

export const getDateHeader = (someDate) => {
    let weekday = new Intl.DateTimeFormat('en', { weekday: 'short' }).format(someDate)
    
    if (isToday(someDate)) {
        weekday = 'Today';
    } else if (isTomorrow(someDate)) {
        weekday = 'Tomorrow';
    }

    const month = new Intl.DateTimeFormat('en', { month: 'short' }).format(someDate)
    const day = new Intl.DateTimeFormat('en', { day: 'numeric' }).format(someDate)

    return `${weekday}, ${month} ${day}`;
}

export const getDuration = (start: Date, end: Date) => {
    const diff = Math.abs(end.valueOf() - start.valueOf());

    let minutes = diff / 1000 / 60;

    if (minutes < 60) {
        return `${minutes} minutes`
    }

    const hour = minutes / 60;

    return `${hour} hr` + (hour > 1 ? 's' : '');
};

export const getFormatedTime = (date: Date) => {

    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    return new Intl.DateTimeFormat('en', { hour: 'numeric', minute: 'numeric' }).format(date)
};
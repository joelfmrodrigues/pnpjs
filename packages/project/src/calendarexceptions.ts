import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { Calendar } from "./calendars";

/**
 * Represents a collection of calendar exceptions
 */
export class CalendarExceptionCollection extends ProjectQueryableCollection {

    /**
    * Gets a calendar exception from the collection of exceptions with the specified object identifier
    *
    * @param id A calendar exception integer identifier
    */
    public getById(id: number): CalendarException {
        const calendarException = new CalendarException(this);
        calendarException.concat(`(${id})`);
        return calendarException;
    }

    /**
     * Adds a calendar exception to the collection of calendar exceptions
     *
     * @param parameters An object that contains information, for example start and finish times and dates, about a new calendar exception
     */
    public async add(parameters: CalendarExceptionCreationInformation): Promise<CommandResult<CalendarException>> {
        this.concat("/add");
        const params: any = {
            "parameters": parameters,
        };
        const data = await this.postCore({ body: jsS(params) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a difference (an exception) from the base calendar
 */
export class CalendarException extends ProjectQueryableInstance {

    /**
     * Gets the calendar that is associated with the exception
     */
    public get calendar(): Calendar {
        return new Calendar(this, "Calendar");
    }

    /**
     * Deletes the CalendarException object
     */
    public delete = this._delete;
}

/**
 * Represents the collection of base calendar exceptions
 */
export class BaseCalendarException extends CalendarException {
}

/**
 * Provides information for the creation of a calendar exception
 */
export interface CalendarExceptionCreationInformation {

    /**
     * Gets or sets the date and time that the calendar exception ends
     */
    Finish?: Date;

    /**
     * Gets or sets the name given for the calendar exception, such as Vacation
     */
    Name?: string;

    /**
     * Gets or sets a mask that represents the days of the week on which the calendar exception is effective
     */
    RecurrenceDays?: CalendarRecurrenceDays;

    /**
     * Gets or sets the interval at which the calendar exception occurs
     */
    RecurrenceFrequency?: number;

    /**
     * Gets or sets the month when setting a yearly recurrence
     */
    RecurrenceMonth?: number;

    /**
     * Gets or sets the day of the month when setting a yearly recurrence
     */
    RecurrenceMonthDay?: number;

    /**
     * Gets or sets the recurrence type for the calendar exception
     */
    RecurrenceType?: CalendarRecurrenceType;

    /**
     * Gets or sets the week number of a monthly occurrence
     */
    RecurrenceWeek?: CalendarRecurrenceWeek;

    /**
     * Gets or sets the minute of the day that the first shift ends
     */
    Shift1Finish?: number;

    /**
     * Gets or sets the minute of the day that the first shift starts
     */
    Shift1Start?: number;

    /**
     * Gets or sets the minute of the day that the second shift ends
     */
    Shift2Finish?: number;

    /**
     * Gets or sets the minute of the day that the second shift starts
     */
    Shift2Start?: number;

    /**
     * Gets or sets the minute of the day that the third shift ends
     */
    Shift3Finish?: number;

    /**
     * Gets or sets the minute of the day that the third shift starts
     */
    Shift3Start?: number;

    /**
     * Gets or sets the minute of the day that the fourth shift ends
     */
    Shift4Finish?: number;

    /**
     * Gets or sets the minute of the day that the fourth shift starts
     */
    Shift4Start?: number;

    /**
     * Gets or sets the minute of the day that the fifth shift ends
     */
    Shift5Finish?: number;

    /**
     * Gets or sets the minute of the day that the fifth shift starts
     */
    Shift5Start?: number;

    /**
     * Gets or sets the date and time that the calendar exception starts
     */
    Start?: Date;
}

/**
 * Represents the days of the week for recurring calendar exceptions
 */
export enum CalendarRecurrenceDays {
    NotSpecified = 0,
    Sunday = 1,
    Monday = 2,
    Tuesday = 4,
    Wednesday = 8,
    Thursday = 16, // 0x00000010
    Friday = 32, // 0x00000020
    Saturday = 64, // 0x00000040
}

/**
 * Specifies the recurrence types for a calendar exception
 */
export enum CalendarRecurrenceType {

    /**
     * Daily recurrence
     */
    Daily,

    /**
     * Daily, with recurrence defined by skipping a set number of days
     */
    DailySkip,

    /**
     * Weekly recurrence
     */
    Weekly,

    /**
     * Monthly recurrence
     */
    Monthly,

    /**
     * Yearly recurrence
     */
    Yearly,
}

/**
 * Specifies one week of a month that is used to setup a schedule
 */
export enum CalendarRecurrenceWeek {

    /**
     * The week number of a monthly occurrence is not specified
     */
    NotSpecified,

    /**
     * First week of a month
     */
    First,

    /**
     * Second week of a month
     */
    Second,

    /**
     * Third week of a month
     */
    Third,

    /**
     * Fourth week of a month
     */
    Fourth,

    /**
     * Last week of a month
     */
    Last,
}

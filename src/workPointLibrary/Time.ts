import { IDatePickerStrings } from 'office-ui-fabric-react';
import * as strings from 'WorkPointStrings';
import { isValidDate } from "./Helper";

/**
* Milliseconds representing common date objects
*/
const millisecondsPerMinute = 1000 * 60;
const millisecondsPerHour = 1000 * 60 * 60;
const millisecondsPerDay = 1000 * 60 * 60 * 24;
const millisecondsPerWeek = 1000 * 60 * 60 * 24 * 7;
const millisecondsPerMonth = 1000 * 60 * 60 * 24 * 30;
const millisecondsPerYear = 1000 * 60 * 60 * 24 * 7 * 365;

/**
* Localizations of common time strings
*/
const minute:string = strings.Minute.toLowerCase();
const minutes:string = strings.Minutes.toLowerCase();

const hour:string = strings.Hour.toLowerCase();
const hours:string = strings.Hours.toLowerCase();

const day:string = strings.Day.toLowerCase();
const days:string = strings.Days.toLowerCase();

const week:string = strings.Week.toLowerCase();
const weeks:string = strings.Weeks.toLowerCase();

const month:string = strings.Month.toLowerCase();
const months:string = strings.Months.toLowerCase();

const year:string = strings.Year.toLowerCase();
const years:string = strings.Years.toLowerCase();

const ago:string = strings.Ago.toLowerCase();

export const getDateTimeFromString = (timeString:string):Date => {
  
  try {
    
    let _time:Date = null;
    
    if (typeof timeString !== "string" || timeString === "") {
      throw "Not a valid candidate for creating a date object.";
    }
    
    _time = new Date(timeString);
    
    if (!isValidDate(_time)) {
      throw "Not a valid date object.";
    }
    
    return _time;
    
  } catch (exception) {
    return null;
  }
  
};

/**
* Prvoide a textual time measuring the time between a DateTime and now.
* 
* TODO: Can we make this via one variable and typeguards?
* 
* @param timeString A string representaion of a date
* @param timeDate Optional Date object to use instead of timeString.
* 
* @returns A string something like this: "Just now"
*/
export const getUserFriendlyTime = (timeString:string, timeDate:Date = null):string => {
  
  try {
    
    let _time:Date = null;
    
    if (timeString === null && isValidDate(timeDate)) {
      _time = timeDate;
    } else if (typeof timeString === "string" && timeString !== "") {
      _time = new Date(timeString);
    }
    
    if (!isValidDate(_time)) {
      throw "Not a valid date object.";
    }
    
    const _now:Date = new Date();
    
    const UTCTime:Date = new Date(Date.UTC(_time.getFullYear(), _time.getMonth(), _time.getDate(), _time.getHours(), _time.getMinutes()));
    const UTCNow:Date = new Date(Date.UTC(_now.getFullYear(), _now.getMonth(), _now.getDate(), _now.getHours(), _now.getMinutes()));
    
    /**
    * Millisecond difference
    */
    const difference:number = UTCNow.getTime() - UTCTime.getTime();
    
    const minuteDifference:number = Math.floor(difference / millisecondsPerMinute);
    const hourDifference:number = Math.floor(difference / millisecondsPerHour);
    const dayDifference:number = Math.floor(difference / millisecondsPerDay);
    const weekDifference:number = Math.floor(difference / millisecondsPerWeek);
    const monthDifference:number = Math.floor(difference / millisecondsPerMonth);
    const yearDifference:number = Math.floor(difference / millisecondsPerYear);
    
    if (minuteDifference < 3) {
      return strings.JustNow;
    }
    
    if (minuteDifference > 3 && minuteDifference < 60) {
      return `${minuteDifference} ${minutes} ${ago}`;
    }
    
    if (hourDifference > 0 && dayDifference < 1) {
      return `${hourDifference} ${hourDifference === 1 ? hour : hours} ${ago}`;
    }
    
    if (dayDifference > 0 && weekDifference < 1) {
      return `${dayDifference} ${dayDifference === 1 ? day : days} ${ago}`;
    }
    
    if (weekDifference > 0 && monthDifference < 1) {
      return `${weekDifference} ${weekDifference === 1 ? week : weeks} ${ago}`;
    }
    
    if (monthDifference > 0 && yearDifference < 1) {
      return `${monthDifference} ${monthDifference === 1 ? month : months} ${ago}`;
    }
    
    if (yearDifference > 0) {
      return `${yearDifference} ${yearDifference === 1 ? year : years} ${ago}`;
    }
    
  } catch (exception) {
    return null;
  }
  
};

export const getAspNetTicksFromDate = (dateTime:Date) => {
  // the number of .net ticks at the unix epoch
  var epochTicks = 621355968000000000;
  
  // there are 10000 .net ticks per millisecond
  var ticksPerMillisecond = 10000;
  
  // calculate the total number of .net ticks for your date
  var totalTicks = epochTicks + (dateTime.getTime() * ticksPerMillisecond);
  
  return totalTicks;
};

export const dayPickerStrings:IDatePickerStrings = {
  months: [
    strings.January,
    strings.February,
    strings.March,
    strings.April,
    strings.May,
    strings.June,
    strings.July,
    strings.August,
    strings.September,
    strings.October,
    strings.November,
    strings.December
  ],

  shortMonths: [
    strings.JanuaryShort,
    strings.FebruaryShort,
    strings.MarchShort,
    strings.AprilShort,
    strings.MayShort,
    strings.JuneShort,
    strings.JulyShort,
    strings.AugustShort,
    strings.SeptemberShort,
    strings.OctoberShort,
    strings.NovemberShort,
    strings.DecemberShort
  ],

  days: [
    strings.Sunday,
    strings.Monday,
    strings.Tuesday,
    strings.Wednesday,
    strings.Thursday,
    strings.Friday,
    strings.Saturday
  ],

  shortDays: [
    strings.SundayShort,
    strings.MondayShort,
    strings.TuesdayShort,
    strings.WednesdayShort,
    strings.ThursdayShort,
    strings.FridayShort,
    strings.SaturdayShort
  ],

  goToToday: strings.GoToToday,
  prevMonthAriaLabel: strings.GoToPreviousMonth,
  nextMonthAriaLabel: strings.GoToNextMonth,
  prevYearAriaLabel: strings.GoToPreviousYear,
  nextYearAriaLabel: strings.GoToNextYear
};
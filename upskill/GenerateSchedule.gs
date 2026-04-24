/**
 * Generates an 80-working-day schedule starting from the given date, skipping weekends,
 * public holidays, and the winter break. Days are grouped into arrays of 5 (Mon–Fri weeks).
 *
 * Also tracks the start and end dates of the two project windows and two hackathon windows
 * based on fixed day-number thresholds.
 *
 * Returns a scheduleData object containing:
 *  - weeks: total number of weeks
 *  - schedule: 2D array of { day, date } objects grouped by week
 *  - proj1StartDate, proj1EndDate, proj2StartDate, proj2EndDate
 *  - hack1StartDate, hack1EndDate, hack2StartDate, hack2EndDate
 */
// Fixed day-number thresholds for project and hackathon windows.
// Defined at module scope so isProjectOrHackathonDay can reference the same values
// without duplicating them.
const PROJ1_START_DAY = 21;
const PROJ1_END_DAY = 25;
const HACK1_START_DAY = 36;
const HACK1_END_DAY = 41;
const PROJ2_START_DAY = 60;
const PROJ2_END_DAY = 75;
const HACK2_START_DAY = 76;
const HACK2_END_DAY = 80;

function generateSchedule(startDate, holidays) {
    let currentDate = new Date(startDate);

    let schedule = [];
    let scheduleLength = 0;
    let week = [];
    let proj1StartDate;
    let proj1EndDate;
    let hack1StartDate;
    let hack1EndDate;
    let proj2StartDate;
    let proj2EndDate;
    let hack2StartDate;
    let hack2EndDate;

    while (scheduleLength < 80) {
        // Skip non-working days without counting them toward the 80-day total.
        if (
            isWeekend(currentDate) ||
            isPublicHoliday(currentDate, holidays) ||
            isWinterBreak(currentDate)
        ) {
            currentDate.setDate(currentDate.getDate() + 1);
            continue;
        }

        scheduleLength++;

        // Normalise to midday to avoid timezone-related date shifts when converting to ISO string.
        currentDate.setHours(12, 0, 0, 0);
        let formattedCurentDate = currentDate.toISOString().split("T")[0];

        let dayOfWeek = currentDate.getDay();

        let dateObj = {
            day: scheduleLength,
            date: formattedCurentDate,
        };
        week.push(dateObj);

        // Once a full Mon–Fri week is accumulated, push it to the schedule and start a new week.
        if (week.length === 5) {
            schedule.push(week);
            week = [];
        }

        currentDate.setDate(currentDate.getDate() + 1);

        // Capture the dates that mark the boundaries of project and hackathon windows.
        if (scheduleLength == PROJ1_START_DAY) proj1StartDate = formattedCurentDate;
        if (scheduleLength == PROJ1_END_DAY) proj1EndDate = formattedCurentDate;
        if (scheduleLength == PROJ2_START_DAY) proj2StartDate = formattedCurentDate;
        if (scheduleLength == PROJ2_END_DAY) proj2EndDate = formattedCurentDate;
        if (scheduleLength == HACK1_START_DAY) hack1StartDate = formattedCurentDate;
        if (scheduleLength == HACK1_END_DAY) hack1EndDate = formattedCurentDate;
        if (scheduleLength == HACK2_START_DAY) hack2StartDate = formattedCurentDate;
        if (scheduleLength == HACK2_END_DAY) hack2EndDate = formattedCurentDate;
    }

    // Push any remaining days that didn't fill a complete week.
    if (week.length > 0) {
        schedule.push(week);
    }

    let scheduleData = {
        weeks: schedule.length,
        schedule: schedule,
        proj1StartDate: proj1StartDate,
        proj1EndDate: proj1EndDate,
        proj2StartDate: proj2StartDate,
        proj2EndDate: proj2EndDate,
        hack1StartDate: hack1StartDate,
        hack1EndDate: hack1EndDate,
        hack2StartDate: hack2StartDate,
        hack2EndDate: hack2EndDate,
    };

    return scheduleData;
}

/**
 * Returns true if the given date falls on a Saturday or Sunday.
 */
function isWeekend(date) {
    const day = date.getDay();
    return day === 0 || day === 6; // Sunday = 0, Saturday = 6
}

/**
 * Returns true if the given date matches any date in the provided holidays array.
 */
function isPublicHoliday(date, holidays) {
    return holidays.some(
        (holiday) =>
            holiday.getFullYear() === date.getFullYear() &&
            holiday.getMonth() === date.getMonth() &&
            holiday.getDate() === date.getDate(),
    );
}

/**
 * Returns true if the given date falls within the winter break: 24 December to 1 January (inclusive).
 */
function isWinterBreak(date) {
    const month = date.getMonth();
    const day = date.getDate();
    return (month === 11 && day >= 24) || (month === 0 && day <= 1);
}

/**
 * Formats a Date object into a long-form locale string, e.g. "Monday, January 1, 2024".
 */
function formatDateLong(date) {
    const options = {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
    };
    return date.toLocaleDateString("en-US", options);
}

/**
 * Returns a hardcoded array of public holiday Date objects for the given location.
 * Supported locations: "IE" (Ireland) or any other value (defaults to England & Wales).
 * Covers 2024–2026.
 */
function getPublicHolidays(location) {
    const irelandHolidays = [
        // 2024
        new Date("2024-01-01"), // New Year's Day
        new Date("2024-02-05"), // St. Brigid's Day
        new Date("2024-03-18"), // St. Patrick's Day (observed)
        new Date("2024-04-01"), // Easter Monday
        new Date("2024-05-06"), // May Day
        new Date("2024-06-03"), // June Bank Holiday
        new Date("2024-08-05"), // August Bank Holiday
        new Date("2024-10-28"), // October Bank Holiday
        new Date("2024-12-25"), // Christmas Day
        new Date("2024-12-26"), // St. Stephen's Day
        // 2025
        new Date("2025-01-01"), // New Year's Day
        new Date("2025-02-03"), // St. Brigid's Day
        new Date("2025-03-17"), // St. Patrick's Day
        new Date("2025-04-21"), // Easter Monday
        new Date("2025-05-05"), // May Day
        new Date("2025-06-02"), // June Bank Holiday
        new Date("2025-08-04"), // August Bank Holiday
        new Date("2025-10-27"), // October Bank Holiday
        new Date("2025-12-25"), // Christmas Day
        new Date("2025-12-26"), // St. Stephen's Day
        // 2026
        new Date("2026-01-01"), // New Year's Day
        new Date("2026-02-02"), // St. Brigid's Day
        new Date("2026-03-17"), // St. Patrick's Day
        new Date("2026-04-06"), // Easter Monday
        new Date("2026-05-04"), // May Day
        new Date("2026-06-01"), // June Bank Holiday
        new Date("2026-08-03"), // August Bank Holiday
        new Date("2026-10-26"), // October Bank Holiday
        new Date("2026-12-25"), // Christmas Day
        new Date("2026-12-28"), // St. Stephen's Day (observed)
    ];

    const englandWalesHolidays = [
        // 2024
        new Date("2024-01-01"), // New Year's Day
        new Date("2024-03-29"), // Good Friday
        new Date("2024-04-01"), // Easter Monday
        new Date("2024-05-06"), // Early May Bank Holiday
        new Date("2024-05-27"), // Spring Bank Holiday
        new Date("2024-08-26"), // Summer Bank Holiday
        new Date("2024-12-25"), // Christmas Day
        new Date("2024-12-26"), // Boxing Day
        // 2025
        new Date("2025-01-01"), // New Year's Day
        new Date("2025-04-18"), // Good Friday
        new Date("2025-04-21"), // Easter Monday
        new Date("2025-05-05"), // Early May Bank Holiday
        new Date("2025-05-26"), // Spring Bank Holiday
        new Date("2025-08-25"), // Summer Bank Holiday
        new Date("2025-12-25"), // Christmas Day
        new Date("2025-12-26"), // Boxing Day
        // 2026
        new Date("2026-01-01"), // New Year's Day
        new Date("2026-04-03"), // Good Friday
        new Date("2026-04-06"), // Easter Monday
        new Date("2026-05-04"), // Early May Bank Holiday
        new Date("2026-05-25"), // Spring Bank Holiday
        new Date("2026-08-31"), // Summer Bank Holiday
        new Date("2026-12-25"), // Christmas Day
        new Date("2026-12-28"), // Boxing Day (observed)
    ];

    return location === "IE" ? irelandHolidays : englandWalesHolidays;
}

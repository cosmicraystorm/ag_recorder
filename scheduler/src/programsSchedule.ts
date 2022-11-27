const ROW_INDEX_MAX = 49;
const TRIGGER_ID_STORE_KEY_PREFIX = "trigger_";
const NOTICE_ADDRESSES_STORE_KEY = "ag_record_start_notice_contacts";

const displayTimeToDatetime = (now: Date, displayTime: string) => {
    const splited = displayTime.split(":");
    let hour = parseInt(splited[0]);
    const minutes = parseInt(splited[1]);
    
    let addDay = 0;
    if (hour >= 24) {
        addDay = 1;
        hour = hour - 24;
    }

    return new Date(now.getFullYear(), now.getMonth(), now.getDate() + addDay, hour, minutes);
}
const getContacts = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    const addresses = [];
    for (let i = 0; i < ROW_INDEX_MAX; i++) {
        const address = sheet.getRange(`R${i + 1}C1`).getDisplayValue();
        if (address === "") {
            break;
        }
        addresses.push(address);
    }
    return addresses;
}

export interface ProgramSchedule {
    startTime: Date;
    endTime: Date;
    name: string;
}


// 60分以上の番組は見たことないので分割処理省略
export const getTodayProgramsSchedule = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    const now = new Date();

    const timeTable = sheet.getRange(`R2C1:R${ROW_INDEX_MAX}C1`).getDisplayValues().map(row => displayTimeToDatetime(now, row[0]));

    const targetScheduleColIndex = now.getDay() + 1 + 1;
    const schedule = sheet.getRange(`R2C${targetScheduleColIndex}:R${ROW_INDEX_MAX}C${targetScheduleColIndex}`).getDisplayValues().map(row => row[0]);

    const programSchedules: ProgramSchedule[] = [];
    let i = 0;
    while (i < schedule.length) {
        if (schedule[i] === "") {
            i++;
            continue;
        }

        const continuousPrograms = [{time: timeTable[i], name: schedule[i]}];
        while (i < schedule.length && schedule[i] === schedule[i + 1]) {
            i++;
            continuousPrograms.push({time: timeTable[i], name: schedule[i]});
        }

        const lastTime = continuousPrograms[continuousPrograms.length - 1].time;
        const endTime = new Date(lastTime.getFullYear(), lastTime.getMonth(), lastTime.getDate(), lastTime.getHours(), lastTime.getMinutes() + 30);
        programSchedules.push({
            startTime: continuousPrograms[0].time,
            endTime: endTime,
            name: continuousPrograms[0].name
        });

        i++;
    }

    return programSchedules;
}

const deleteTrigger = (trigger: GoogleAppsScript.Script.Trigger) => ScriptApp.deleteTrigger(trigger);

function triggerRecordStart(e: any) {
    const currentTrigger = ScriptApp.getProjectTriggers().filter(trigger => trigger.getUniqueId() == e.triggerUid)[0];
    const store = PropertiesService.getDocumentProperties();
    const dataStr = store.getProperty(TRIGGER_ID_STORE_KEY_PREFIX + e.triggerUid);
    if (!dataStr) {
        deleteTrigger(currentTrigger);
        return
    }
    const schedule: ProgramSchedule = JSON.parse(dataStr);
    store.deleteProperty(TRIGGER_ID_STORE_KEY_PREFIX + e.triggerUid);

    const contacts: string[] = JSON.parse(store.getProperty(NOTICE_ADDRESSES_STORE_KEY)!);
    GmailApp.sendEmail(
        contacts[0],
        `start record ${schedule.name}`,
        `start record ${schedule.name} by ag_record_schedule`,
        {
            cc: contacts.length > 1 ? contacts.slice(1).join(",") : undefined,
            from: contacts[0]   // noreplyは個人アカウントで使えないので先頭をfromにする
        }
    );
    // Todo: request start record to pub/sub

    // delete self
    deleteTrigger(currentTrigger);
}

function triggerDayStart() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const schedules = getTodayProgramsSchedule(spreadsheet.getSheetByName("schedule")!);
    const store = PropertiesService.getDocumentProperties();
    for (let i = 0; i < schedules.length; i++) {
        const triggerFireTime = new Date(schedules[i].startTime.getTime() - (30 * 1000));   // 30sec早く起動させて各処理やコンテナ起動の時間を稼ぐ
        const trigger = ScriptApp.newTrigger("triggerRecordStart").timeBased().at(triggerFireTime).create();
        const trigerId = trigger.getUniqueId();
        store.setProperty(TRIGGER_ID_STORE_KEY_PREFIX + trigerId, JSON.stringify(schedules[i]));
    }

    // read contacts
    if (store.getKeys().indexOf(NOTICE_ADDRESSES_STORE_KEY) >= 0) {
        store.deleteProperty(NOTICE_ADDRESSES_STORE_KEY);
    }
    store.setProperty(NOTICE_ADDRESSES_STORE_KEY, JSON.stringify(getContacts(spreadsheet.getSheetByName("contact")!)));
}

function test() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("contact");
    console.log(getContacts(sheet!));
}
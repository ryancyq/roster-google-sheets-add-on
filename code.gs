/*
 * Helper function to get the default configurations
 */
function geDefaultConfig() {
    return {
        is_initialized: false,
        lookup: {
            sheet_name: 'Sheet1',
            range: {
                person_name: 'A:A',
                timeslot: 'B:B',
                timestamp: 'C:C'
            }
        },
        fillup: {
            sheet_name: 'Sheet2',
            range: {
                person_name: 'A:A',
                timetable_weekly: 'B:H',
                timestamp: 'I:I'
            },
            schedule_weekly: [1, 2, 3, 4, 5, 6, 7]
        },
        data_retention: {
            expiry_days: -1
        }
    }
};

/*
 * Helper function to read the configurations from properties service (e.g User/Document/App)
 */
function readConfigFromProperties(props) {
    var config = geDefaultConfig();
    try {
        config.is_initialized = props.getProperty('IS_INITIALIZED');

        config.lookup.sheet_name = props.getProperty('LOOKUP_SHEET_NAME');
        config.lookup.range.person_name = props.getProperty('LOOKUP_RANGE_PERSON_NAME');
        config.lookup.range.timeslot = props.getProperty('LOOKUP_RANGE_TIMESLOT');
        config.lookup.range.timestamp = props.getProperty('LOOKUP_RANGE_TIMESTAMP');

        config.fillup.sheet_name = props.getProperty('FILLUP_SHEET_NAME');
        config.fillup.range.person_name = props.getProperty('FILLUP_RANGE_PERSON_NAME');
        config.fillup.range.timetable_weekly = props.getProperty('FILLUP_RANGE_TIMETABLE_WEEKLY');
        config.fillup.range.timestamp = props.getProperty('FILLUP_RANGE_TIMESTAMP');

        config.data_retention.expiry_days = props.getProperty('DATE_RETENTION_EXPIRY_DAYS');

    } catch (e) {
        throw "Unable to read config for the sheet"
    }
}

/*
 * Helper function to save the configurations to properties service (e.g User/Document/App)
 */
function saveConfigToProperties(config, props) {
    try {
        props.setProperties({
            // 'IS_INITIALIZED' : config.is_initialized,

            'LOOKUP_SHEET_NAME': config.lookup.sheet_name,
            'LOOKUP_RANGE_PERSON_NAME': config.lookup.range.person_name,
            'LOOKUP_RANGE_TIMESLOT': config.lookup.range.timeslot,
            'LOOKUP_RANGE_TIMESTAMP': config.lookup.range.timestamp,

            'FILLUP_SHEET_NAME': config.fillup.sheet_name,
            'FILLUP_RANGE_PERSON_NAME': config.fillup.range.person_name,
            'FILLUP_RANGE_TIMETABLE_WEEKLY': config.fillup.range.timetable_weekly,
            'FILLUP_RANGE_TIMESTAMP': config.fillup.range.timestamp,

            'DATE_RETENTION_EXPIRY_DAYS': config.data_retention.expiry_days,
        });

    } catch (e) {
        throw "Unable to save config for the sheet"
    }
}

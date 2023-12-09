// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const https = require('https');
const { ActivityHandler, MessageFactory } = require('botbuilder');

class EchoBot extends ActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            const owners = await getOwnerOfVehicle(context.activity.text);
            let replyText = `You requested the owner of the vehicle number: ${ context.activity.text }\r\n`;
            if (!owners || !owners.length) {
                replyText += 'No owner found for the given vehicle number';
            } else {
                replyText += 'Following owner(s) matched the given vehicle number\r\n';
                owners.forEach(owner => {
                    replyText += `Name: ${ owner[0] }, Vehicle Number(s): ${ owner[1] }, Phone: ${ owner[2] }\r\n`;
                });
            }
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

async function getOwnerOfVehicle(vehicleNo) {
    return new Promise((resolve) => {
        let data = '';

        https.get('https://docs.google.com/spreadsheets/d/15pc99IZYpcQvxlwF5mWGe0ECd6Q00Jc_7ThqOH5fUpw/export?format=csv', res => {
            https.get(res.headers.location, res => {
                res.on('data', chunk => {
                    data += chunk;
                });
                res.on('end', () => {
                    data = CSVToArray(data.toString('utf8'));
                    const result = data.filter((value, index, obj) => {
                        return value[1].includes(vehicleNo);
                    });
                    resolve(result);
                });
            }).on('error', err => {
                console.log('Error: ', err.message);
            });
        }).on('error', err => {
            console.log('Error: ', err.message);
        });
    });
}

/**
 * CSVToArray parses any String of Data including '\r' '\n' characters,
 * and returns an array with the rows of data.
 * @param {String} csv - the CSV string you need to parse
 * @param {String} delimiter - the delimeter used to separate fields of data
 * @returns {Array} rows - rows of CSV where first row are column headers
 */
function CSVToArray(csv, delimiter) {
    delimiter = (delimiter || ','); // user-supplied delimeter or default comma

    var pattern = new RegExp( // regular expression to parse the CSV values.
        ( // Delimiters:
            '(\\' + delimiter + '|\\r?\\n|\\r|^)' +
        // Quoted fields.
        '(?:"([^"]*(?:""[^"]*)*)"|' +
        // Standard fields.
        '([^"\\' + delimiter + '\\r\\n]*))'
        ), 'gi'
    );

    var rows = [[]]; // array to hold our data. First row is column headers.
    // array to hold our individual pattern matching groups:
    var matches = false; // false if we don't find any matches
    // Loop until we no longer find a regular expression match
    while (matches = pattern.exec(csv)) {
        var matchedDelimiter = matches[1]; // Get the matched delimiter
        // Check if the delimiter has a length (and is not the start of string)
        // and if it matches field delimiter. If not, it is a row delimiter.
        if (matchedDelimiter.length && matchedDelimiter !== delimiter) {
            // Since this is a new row of data, add an empty row to the array.
            rows.push([]);
        }
        var matchedValue;
        // Once we have eliminated the delimiter, check to see
        // what kind of value was captured (quoted or unquoted):
        if (matches[2]) { // found quoted value. unescape any double quotes.
            matchedValue = matches[2].replace(
                new RegExp('""', 'g'), '"'
            );
        } else { // found a non-quoted value
            matchedValue = matches[3];
        }
        // Now that we have our value string, let's add
        // it to the data array.
        rows[rows.length - 1].push(matchedValue);
    }
    return rows; // Return the parsed data Array
}

module.exports.EchoBot = EchoBot;

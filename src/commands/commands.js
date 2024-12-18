const { salaryData } = require('./salary');

let mailboxItem;

// To check if the manifest.xml has any errors run:
// `npm run validate`

// To test the add-in, run `npm run dev-server` and then side-load the add-in in Outlook. Each time this file is saved, it will recompile and make the new changes available in the add-in.
// To enable developer extras in Outlook, run the following command in the terminal. This will open a dev tools window to see the console logs when the add-in is running.
// `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

// To disable the developer extras, run the following command in the terminal.
// `defaults delete com.microsoft.Outlook OfficeWebAddinDeveloperExtras`

Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

function insertPoints(event) {
    // Get HTML body from the client.
    mailboxItem.body.getAsync("html", { asyncContext: event }, function (getBodyResult) {
        if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
            getTotalDuration(function (duration) {
                console.log("Meeting duration: " + duration + " hours");
                getAttendees(function (attendees) {
                    if (attendees) {
                        resolveJobTitles(attendees, function (jobTitles) {
                            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                                const totalCost = calculateMeetingCost(jobTitles, duration);
                                displayInfobar(`Total cost of the meeting: $${totalCost}`);
                                // For Debugging
                                // const emailsWithJobTitles = jobTitles.map(item => `${item.email} (${item.jobTitle})`).join(", ");
                                // updateBody(getBodyResult.asyncContext, getBodyResult.value, emailsWithJobTitles);
                                // End Debugging
                                getBodyResult.asyncContext.completed({ allowEvent: false }); // This will end the add-in run
                            } else {
                                console.error("Failed to get HTML body.");
                                getBodyResult.asyncContext.completed({ allowEvent: false });
                            }
                        });
                    } else {
                        console.error("Failed to get attendees.");
                        getBodyResult.asyncContext.completed({ allowEvent: false });
                    }
                });
            });
        }
    });
}

function getAttendees(callback) {
    let attendees = [];

    mailboxItem.requiredAttendees.getAsync(function (requiredResult) {
        if (requiredResult.status === Office.AsyncResultStatus.Succeeded) {
            attendees = requiredResult.value.map(attendee => attendee.emailAddress);

            mailboxItem.optionalAttendees.getAsync(function (optionalResult) {
                if (optionalResult.status === Office.AsyncResultStatus.Succeeded) {
                    attendees = attendees.concat(optionalResult.value.map(attendee => attendee.emailAddress));
                    
                    // Expand distribution lists
                    expandDistributionLists(attendees, callback);
                } else {
                    console.error("Failed to get optional attendees.");
                    callback(null);
                }
            });
        } else {
            console.error("Failed to get required attendees.");
            callback(null);
        }
    });
}

function expandDistributionLists(attendees, callback) {
    let expandedAttendees = [];
    let pendingRequests = attendees.length;

    attendees.forEach(email => {
        // DLs will not have an underscore as part of the email
        // If the email contains an underscore, it is not a DL
        if (email.includes("_")) {
            // Do not add duplicate entries
            if (!expandedAttendees.includes(email)) {
                expandedAttendees.push(email);
            }
            pendingRequests--;
            if (pendingRequests === 0) {
                callback(expandedAttendees);
            }
            return;
        }
        Office.context.mailbox.makeEwsRequestAsync(
            `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Header>
                    <t:RequestServerVersion Version="Exchange2013" />
                </soap:Header>
                <soap:Body>
                    <m:ExpandDL>
                        <m:Mailbox>
                            <t:EmailAddress>${email}</t:EmailAddress>
                        </m:Mailbox>
                    </m:ExpandDL>
                </soap:Body>
            </soap:Envelope>`,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
                    const members = xmlDoc.getElementsByTagName("t:Mailbox");

                    // Not a distribution list if there are no members
                    if (members.length === 0) {
                        // Do not add duplicate entries
                        if (!expandedAttendees.includes(email)) {
                            expandedAttendees.push(email);
                        }
                    } else {
                        for (let i = 0; i < members.length; i++) {
                            const memberEmail = members[i].getElementsByTagName("t:EmailAddress")[0].textContent;
                            // Do not add duplicate entries
                            if (!expandedAttendees.includes(memberEmail)) {
                                expandedAttendees.push(memberEmail);
                            }
                        }
                    }
                } else {
                    console.error("Failed to expand distribution list: " + email);
                    expandedAttendees.push(email); // If expansion fails, keep the original email
                }

                pendingRequests--;
                if (pendingRequests === 0) {
                    console.log(expandedAttendees)
                    callback(expandedAttendees);
                }
            }
        );
    });
}

function resolveJobTitles(emails, callback) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let jobTitles = [];
            let pendingRequests = emails.length;

            emails.forEach(email => {
                Office.context.mailbox.makeEwsRequestAsync(
                    `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Header>
                            <t:RequestServerVersion Version="Exchange2013" />
                        </soap:Header>
                        <soap:Body>
                            <m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">
                                <m:UnresolvedEntry>${email}</m:UnresolvedEntry>
                            </m:ResolveNames>
                        </soap:Body>
                    </soap:Envelope>`,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(asyncResult.value, "text/xml");
                            const jobTitle = xmlDoc.getElementsByTagName("t:JobTitle")[0]?.textContent;

                            if (jobTitle) {
                                jobTitles.push({ email, jobTitle });
                            } else {
                                jobTitles.push({ email, jobTitle: "Not found" });
                            }
                        } else {
                            jobTitles.push({ email, jobTitle: "Request failed" });
                        }

                        pendingRequests--;
                        if (pendingRequests === 0) {
                            callback(jobTitles);
                        }
                    }
                );
            });
        } else {
            console.error("Failed to get callback token:" + result.error.code + " " + result.error.name + " " + result.error.message);
            callback(null);
        }
    });
}

function displayInfobar(message) {
    const key = "meetingcost"
    console.log("Displaying infobar: " + message);
    Office.context.mailbox.item.notificationMessages.addAsync(key, {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "Icon.16x16",
        persistent: true
    }, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
            // We can only use `addAsync` once per key, so we need to replace the infobar if it already exists.
            // The most likely reason for failing to add is this, so we will assume that is always the case
            console.error("Failed to add infobar, try to replace instead: " + result.error.message);
            Office.context.mailbox.item.notificationMessages.replaceAsync(key, {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.16x16",
                persistent: true
            }, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to replace infobar: " + result.error.message);
                }
            });
        }
    });

}

function getTotalDuration(callback) {
    const start = mailboxItem.start;
    const end = mailboxItem.end;

    start.getAsync(function (startResult) {
        if (startResult.status === Office.AsyncResultStatus.Succeeded) {
            end.getAsync(function (endResult) {
                if (endResult.status === Office.AsyncResultStatus.Succeeded) {
                    const startTime = new Date(startResult.value);
                    const endTime = new Date(endResult.value);
                    const duration = (endTime - startTime) / (1000 * 60 * 60); // Convert milliseconds to hours
                    callback(duration);
                } else {
                    console.error("Failed to get end time.");
                    callback(null);
                }
            });
        } else {
            console.error("Failed to get start time.");
            callback(null);
        }
    });
}

function calculateMeetingCost(jobTitles, duration) {
    let totalCost = 0;

    jobTitles.forEach(item => {
        console.log(`Email: ${item.email}, Job Title: ${item.jobTitle}`);
        console.log(salaryData);
        console.log(`Salary: ${salaryData[item.jobTitle] || salaryData["Default"]}`);
        const salary = salaryData[item.jobTitle] || salaryData["Default"];
        const hourlyRate = salary / 1950; // Assuming 1950 working hours in a year (37.5 / week)
        totalCost += hourlyRate * duration;
    });

    return totalCost.toFixed(2); // Return cost rounded to 2 decimal places
}

// Register the functions.
Office.actions.associate("insertPoints", insertPoints);
Office.actions.associate("attendeeChanged", attendeeChanged);
Office.actions.associate("durationChanged", durationChanged);

function attendeeChanged(event) {
    console.log(`Event: ${event.type}`);
    insertPoints(null);
}

function durationChanged(event) {
    console.log(`Event: ${event.type}`);
    insertPoints(null);
}

// Use this for debugging to display data in the message
function updateBody(event, existingBody, annotation) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + annotation,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set HTML body.");
            }
        }
    );
}

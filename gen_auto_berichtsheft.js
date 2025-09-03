import fs from 'fs';
import { Document, Packer, Paragraph, TextRun, LevelFormat, AlignmentType, convertInchesToTwip } from "docx";
import { WebUntis } from "webuntis";
import dotenv from "dotenv";
import { get } from 'http';
import readline from "readline";
import path from "path";
import cliProgress from 'cli-progress';

dotenv.config();

const DEBUG = process.env.DEBUG === "true" || process.env.DEBUG === "1";

// Configuration for WebUntis login
const UNTIS_SCHOOL = process.env.UNTIS_SCHOOL;
const UNTIS_USERNAME = process.env.UNTIS_USERNAME;
const UNTIS_PASSWORD = process.env.UNTIS_PASSWORD;
const UNTIS_SERVER = process.env.UNTIS_SERVER;

const formatLocalDateTime = (date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0"); // Months are 0-based
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");

    return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
};

async function loginAndGetCookie() {
    const loginUrl = `https://${UNTIS_SERVER}/WebUntis/jsonrpc.do?school=${UNTIS_SCHOOL}`;
    const loginPayload = {
        id: "login",
        method: "authenticate",
        params: {
            user: UNTIS_USERNAME,
            password: UNTIS_PASSWORD,
        },
        jsonrpc: "2.0",
    };

    const loginResponse = await fetch(loginUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify(loginPayload),
    });

    if (!loginResponse.ok) {
        throw new Error(`Failed to login: ${loginResponse.status} ${loginResponse.statusText}`);
    }

    // Extract the Set-Cookie header
    const setCookieHeader = loginResponse.headers.get("set-cookie");
    if (!setCookieHeader) {
        throw new Error("No Set-Cookie header found in the login response.");
    }

    // Combine all cookies into a single string (if multiple cookies exist)
    const cookie = setCookieHeader.split(",").map(cookie => cookie.split(";")[0]).join("; ");
    if (DEBUG) {
        console.log("Extracted Cookie:", cookie);
    }

    return cookie;
}

// Function to prompt the user for input
function promptUserForTimeframe() {
    return new Promise((resolve, reject) => {
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout,
        });

        rl.question("Enter the start date (YYYY-MM-DD): ", (startDateInput) => {
            rl.question("Enter the end date (YYYY-MM-DD): ", (endDateInput) => {
                rl.close();

                // Validate the input dates
                const startDate = new Date(startDateInput);
                const endDate = new Date(endDateInput);

                if (isNaN(startDate) || isNaN(endDate)) {
                    return reject(new Error("Invalid date format. Please use an valid date as YYYY-MM-DD."));
                }

                if (startDate > endDate) {
                    return reject(new Error("Start date must be earlier than or equal to the end date."));
                }

                resolve({ startDate, endDate });
            });
        });
    });
}

// Fetches the teaching contents for each lesson and builds the Word document
async function fetchTeachingContent(startDate, endDate) {
    try {
        // Login and get the cookie
        const cookie = await loginAndGetCookie();
        if (DEBUG) {
            console.log("Using Cookie:", cookie);
        }
        // Extract tenant-id from the cookie -> API Request requires it
        const tenantIdMatch = cookie.match(/Tenant-Id="([^"]+)"/);
        if (!tenantIdMatch || !tenantIdMatch[1]) {
            throw new Error("Tenant-Id not found in the cookie.");
        }
        const tenantId = tenantIdMatch[1];
        if (DEBUG) {
            console.log("Extracted Tenant-Id:", tenantId);
        }
        // Get API Token
        const tokenResponse = await fetch("https://erato.webuntis.com/WebUntis/api/token/new", {
            method: "GET",
            headers: {
                Authorization: `Basic ${Buffer.from(`${UNTIS_USERNAME}:${UNTIS_PASSWORD}`).toString("base64")}`,
                "Content-Type": "application/json",
                Accept: "application/json",
                Cookie: cookie,
            },
        });

        if (!tokenResponse.ok) {
            throw new Error(`Failed to fetch API token: ${tokenResponse.statusText}`);
        }

        // Read the token as plain text (WebUntis API returns a token in plain text)
        const authToken = await tokenResponse.text();
        if (DEBUG) {
            console.log("API-Token Response:", authToken);
        } else {
            console.log("API-Token was successfully fetched.");
        }
        // Ensure the token is not empty
        if (!authToken || authToken.trim() === "") {
            throw new Error("API token is empty or invalid.");
        }

        const paragraphs = [];

        // Add the main title to the word
        paragraphs.push(
            new Paragraph({
                text: "Berichtsheft",
                heading: "Title",
                alignment: "center",
            })
        );

        // Add the calendar week (KW)
        const weekNumber = getCalendarWeek(startDate);
        paragraphs.push(
            new Paragraph({
                text: `KW${weekNumber}/${startDate.getFullYear()}`,
                alignment: "center",
            })
        );

        // Calculate total days to process
        const totalDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        const progressBar = new cliProgress.SingleBar({
            format: 'Progress: |{bar}| {percentage}% | {value}/{total} days',
            barCompleteChar: '\u2588',
            barIncompleteChar: '\u2591',
            hideCursor: true
        }, cliProgress.Presets.shades_classic);

        console.log("Fetching and processing lessons for the selected date range...");
        progressBar.start(totalDays, 0);

        let currentDate = new Date(startDate);
        let dayCount = 0;
        while (currentDate <= endDate) {
            const formattedDate = formatLocalDateTime(currentDate).split("T")[0]; // Format as YYYY-MM-DD
            if (DEBUG) {
                console.log(`Fetching lessons for date: ${formattedDate}`);
            }
            // Fetch all lessons for the current date
            const lessons = []; // Collect all lessons for the day
            for (let hour = 7; hour <= 18; hour++) { // Iterate over hours (7:00 to 18:00)
                const startDateTime = new Date(currentDate);
                startDateTime.setHours(hour, 0, 0); // Start time: hour:00

                const endDateTime = new Date(currentDate);
                endDateTime.setHours(hour + 1, 0, 0); // End time: hour+1:00

                lessons.push({
                    elementId: 36686,
                    elementType: 5,
                    startDateTime: formatLocalDateTime(startDateTime),
                    endDateTime: formatLocalDateTime(endDateTime),
                });
            }

            // Add a heading for the current weekday and date
            paragraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `${currentDate.toLocaleDateString("de-DE", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}:`,
                            underline: {},
                        }),
                    ],
                    heading: "Heading2",
                })
            );

            // Process each lesson
            const daySubjects = {}; // Group subjects and their teaching contents
            for (const lesson of lessons) {
                try {
                    // Construct the URL
                    const url = `https://erato.webuntis.com/WebUntis/api/rest/view/v2/calendar-entry/detail?elementId=${lesson.elementId}&elementType=${lesson.elementType}&endDateTime=${encodeURIComponent(
                        lesson.endDateTime
                    )}&homeworkOption=DUE&startDateTime=${encodeURIComponent(lesson.startDateTime)}`;

                    if (DEBUG) {
                        console.log("API-Request URL:", url);
                    }

                    // Make the request
                    const teachingContentResponse = await fetch(url, {
                        method: "GET",
                        headers: {
                            Authorization: `Bearer ${authToken}`,
                            "Content-Type": "application/json",
                            Accept: "application/json",
                            Cookie: cookie,
                            "Tenant-Id": tenantId,
                        },
                    });

                    if (!teachingContentResponse.ok) {
                        throw new Error(`Failed to fetch teaching content: ${teachingContentResponse.status} ${teachingContentResponse.statusText}`);
                    }

                    const teachingContentData = await teachingContentResponse.json();

                    const calendarEntries = teachingContentData.calendarEntries || [];
                    for (const entry of calendarEntries) {
                        if (DEBUG) {
                            console.log("Teaching Content Response:", JSON.stringify(teachingContentData, null, 2));
                        }

                        if (!entry.subject?.longName) continue; // Skip entries without a subject

                        // Adjust the subject name if it matches the specific longName
                        let subject = entry.subject.longName;
                        if (subject === "LF 8: Daten systemübergreifend bereitstellen / OOP (SI+IT nur Grundlagen) in Python/Java/ Datenbanke") {
                            subject = "LF 8: Daten systemübergreifend bereitstellen";
                        }

                        // Check if the lesson is canceled
                        if (entry.status === "CANCELLED") {
                            const startTime = new Date(entry.startDateTime).toLocaleTimeString("de-DE", {
                                hour: "2-digit",
                                minute: "2-digit",
                            });
                            const endTime = new Date(entry.endDateTime).toLocaleTimeString("de-DE", {
                                hour: "2-digit",
                                minute: "2-digit",
                            });
                            const canceledText = `Entfallen (${startTime} - ${endTime})`;

                            // Group by subject
                            if (!daySubjects[subject]) {
                                daySubjects[subject] = new Set();
                            }
                            daySubjects[subject].add(canceledText);
                            continue;
                        }

                        // Check if teaching content is provided
                        const teachingContent =
                            entry.teachingContent ||
                            "Kein Lehrstoff durch die Lehrkraft angegeben! Bitte manuell eintragen!";

                        // Group by subject
                        if (!daySubjects[subject]) {
                            daySubjects[subject] = new Set();
                        }
                        daySubjects[subject].add(teachingContent);
                    }
                } catch (lessonError) {
                    console.error(`Error fetching teaching content for lesson: ${lessonError.message}`);
                }
            }

            // Add subjects and their teaching contents to the document
            for (const [subject, contents] of Object.entries(daySubjects)) {
                paragraphs.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: subject,
                                color: "156082", // Set a custom color (hexadecimal)
                            }),
                        ],
                        heading: "Heading3",
                    })
                );

                // Replace the bullet point paragraph creation with this:
                for (const content of contents) {
                    paragraphs.push(
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: content,
                                    color: content.includes("Kein Lehrstoff")
                                        ? "FF0000" // Red text for missing teaching content
                                        : "000000", // Black text for normal content
                                }),
                            ],
                            numbering: {
                                reference: "bullet-points",
                                level: 0,
                            },
                        })
                    );
                }
            }

            // Add an empty line between the weekdays
            paragraphs.push(
                new Paragraph({
                    text: "",
                })
            );

            // Move to the next day
            currentDate.setDate(currentDate.getDate() + 1);
            dayCount++;
            progressBar.update(dayCount);
        }

        progressBar.stop();

        // Generate Word Document
        if (paragraphs.length === 0) {
            throw new Error("No valid teaching content to add to the document.");
        }

        // And update your Document creation to include proper numbering definition:
        const doc = new Document({
            creator: "Fabian",
            title: "TeachingContentOverview",
            description: "A document containing teaching content fetched from WebUntis.",
            numbering: {
                config: [
                    {
                        reference: "bullet-points",
                        levels: [
                            {
                                level: 0,
                                format: LevelFormat.BULLET,
                                text: "•",
                                alignment: AlignmentType.LEFT,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(0.5) },
                                    },
                                },
                            },
                        ],
                    },
                ],
            },
            sections: [
                {
                    properties: {},
                    children: paragraphs,
                },
            ],
        });

        const buffer = await Packer.toBuffer(doc);

        let outputFilename = process.env.OUTPUT_FILENAME && process.env.OUTPUT_FILENAME.trim() !== ""
            ? process.env.OUTPUT_FILENAME.trim()
            : "TeachingContentOverview.docx";
        if (!outputFilename.toLowerCase().endsWith(".docx")) {
            outputFilename += ".docx";
        }

        // Use DOCX_PATH from env, fallback to script dir
        const outputDir = process.env.DOCX_PATH && process.env.DOCX_PATH.trim() !== ""
            ? process.env.DOCX_PATH.trim()
            : __dirname;
        const outputPath = path.isAbsolute(outputFilename)
            ? outputFilename
            : path.join(outputDir, outputFilename);

        fs.writeFileSync(outputPath, buffer);
        console.log(`Teaching content exported successfully to ${outputPath}`);
    } catch (error) {
        console.error("An error occurred:", error);
    }
}

// Helper function to calculate the calendar week (KW)
function getCalendarWeek(date) {
    const target = new Date(date.valueOf());
    const dayNr = (date.getDay() + 6) % 7; // Make Monday the first day of the week
    target.setDate(target.getDate() - dayNr + 3); // Thursday is in the same week
    const firstThursday = new Date(target.getFullYear(), 0, 4);
    const weekNumber = Math.ceil(((target - firstThursday) / 86400000 + firstThursday.getDay() + 1) / 7);
    return weekNumber;
}

function displayStartupScreen() {
    console.clear();
    console.log(`
    ============================================================================
                                Berichtsheft Generator
    ============================================================================
    
    Welcome to the Berichtsheft Generator!
    This script fetches teaching contents from WebUntis
    and generates a Word document with the data.

    Usage:
    - Enter the start and end dates when prompted.
    - Ensure your WebUntis credentials are set in the .env file.

    Features:
    - Teaching content grouped by subject and day.
    - Automatically formatted Word document.
    - Debug mode for detailed logs (set DEBUG=true in .env).

    ==============================================================================
    `);

    // Add a red warning message
    console.log("\x1b[31m%s\x1b[0m", "IMPORTANT: Please check the output for undocumented teaching contents before using the generated Word document!");
    console.log("\x1b[31m%s\x1b[0m", "Ensure all teaching contents are properly documented.");
    console.log("\n");
}

async function main() {
    try {
        displayStartupScreen();
        // Prompt the user for the timeframe
        const { startDate, endDate } = await promptUserForTimeframe();
        console.log(`Fetching data from ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}...`);

        // Pass the dates to the fetching function
        await fetchTeachingContent(startDate, endDate);
    } catch (error) {
        console.error("Error:", error.message);
    }
}

main();
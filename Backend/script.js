const express = require('express');
const bodyParser = require('body-parser');
const puppeteer = require('puppeteer');
const { App } = require('@slack/bolt');
require('dotenv').config();
const fs = require('fs');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const Airtable = require('airtable');
const app1 = express();
const cors = require('cors');
const http = require('http').createServer(app1);
const io = require('socket.io')(http, {
    cors: {
        origin: '*',
    }
});

app1.use(bodyParser.json());
app1.use(bodyParser.urlencoded({ extended: true }));

app1.use(cors());

// Emit socket events for logging messages
function logMessage(message) {
    io.emit('log', message);
}

// Function to create the output directory if it doesn't exist
function createOutputDirectory() {
    const directory = 'Output';

    if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory);
    }
}


async function saveUsers(slackBotToken, signingSecret, airtableApiKey, airtableBaseId, airtableTableName) {
    console.log(slackBotToken, signingSecret, airtableApiKey, airtableBaseId, airtableTableName)
    console.log('Collecting all members in a workspace of Slack...');
    logMessage('Collecting all members in a workspace of Slack...');
    try {
        // Initialize the Bolt app with the provided credentials
        const app = new App({
            token: slackBotToken,
            signingSecret: signingSecret,
        });

        const result = await app.client.users.list({
            token: slackBotToken,
        });

        const teamName = await app.client.team.info({
            token: slackBotToken,
        });

        const airtableBase = new Airtable({ apiKey: airtableApiKey }).base(airtableBaseId);
        const airtableTable = airtableTableName;

        console.log(teamName.team.name);
        createOutputDirectory();
        fs.writeFileSync('Output/members.json', JSON.stringify(result.members, null, 2));
        console.log('Members details collected successfully!');
        logMessage('Members details collected successfully!');
        await scrapeAllMembers(result.members, teamName.team.name, airtableBase, airtableTable);
    } catch (error) {
        console.error('An error occurred while saving users:', error);
        logMessage('An error occurred while saving users:' + error);
        throw new Error(error);
    }
};

const fuzzyExtractOrganization = (email) => {
    const emailParts = email.split('@');
    const domain = emailParts[emailParts.length - 1];
    const domainParts = domain.split('.');

    let organization = '';
    for (let i = 0; i < domainParts.length; i++) {
        const part = domainParts[i];
        if (part !== 'edu' && part !== 'com' && part !== 'net' && part !== 'org') {
            organization += part.charAt(0).toUpperCase() + part.slice(1) + ' ';
        } else {
            break;
        }
    }

    if (organization === 'Gmail') {
        return null;
    } else if (organization === 'Outlook') {
        return null;
    }

    return organization.trim().length > 0 ? organization.trim() : null;
};

async function getLinkedInProfile(member, browser, teamName) {
    const page = await browser.newPage();

    try {
        // Generate the search query including the extracted organization name 
        const orgName = fuzzyExtractOrganization(member.profile.email);
        console.log(`Extracted organization name: ${orgName}`);

        // Generate the search query including the extracted organization name and team name
        let searchQuery = `${member.real_name} site:linkedin.com/in/`;
        if (orgName) {
            searchQuery += ` (${orgName} OR ${teamName})`;
        } else {
            searchQuery += ` ${teamName}`;
        }

        console.log(`Search query: ${searchQuery}`);
        // Perform the search on Google
        await page.goto(`https://www.google.com/search?q=${encodeURIComponent(searchQuery)}`);

        // Wait for the search results to load
        const searchResults = await page.$$('.g');
        if (searchResults.length === 0) {
            console.log('No search results found. Performing an alternative search...');
            logMessage('No search results found. Performing an alternative search...');
            searchQuery = `${member.real_name} site:linkedin.com/in/`;
            await page.goto(`https://www.google.com/search?q=${encodeURIComponent(searchQuery)}`);
            await page.waitForSelector('.g');
        }

        // Extract the LinkedIn profile URL from the search results
        const linkedInUrl = await page.evaluate(() => {
            const searchResult = document.querySelector('.g a[href*="linkedin.com/in/"]');
            return searchResult ? searchResult.href : 'Not found';
        });

        const result = {
            name: member.real_name,
            linkedIn: linkedInUrl,
        };

        return result;
    } catch (error) {
        console.error(`An error occurred while processing user ${member.real_name}:`, error);
        logMessage(`An error occurred while processing user ${member.real_name}:`, error);
        return { name: member.real_name, linkedIn: 'Error' };
    } finally {
        await page.close();
    }
}


async function scrapeAllMembers(members, teamName, airtableBase, airtableTable) {
    const browser = await puppeteer.launch({ headless: false });
    try {
        const data = [];
        const onlyMembers = members.filter(
            (member) => member.is_bot === false && member.real_name !== undefined && !member.real_name.toLowerCase().includes('bot')
        );

        console.log(`Processing ${onlyMembers.length} members...`);
        logMessage(`Processing ${onlyMembers.length} members...`);

        for (let i = 0; i < onlyMembers.length; i++) {
            const member = onlyMembers[i];
            console.log(`Processing user ${member.real_name} (${i + 1}/${onlyMembers.length})...`);
            logMessage(`Processing user ${member.real_name} (${i + 1}/${onlyMembers.length})...`);
            const result = await getLinkedInProfile(member, browser, teamName);
            data.push({
                name: member.real_name,
                description: member.profile.title,
                contact: member.profile.email,
                picture: member.profile.image_512,
                linkedIn: result.linkedIn,
            });
        }

        // Save data to Airtable
        for (const record of data) {
            try {
                await new Promise((resolve, reject) => {
                    airtableBase(airtableTable).create(
                        [
                            {
                                fields: {
                                    Name: record.name,
                                    Description: record.description,
                                    Contact: record.contact,
                                    Picture: record.picture,
                                    LinkedIn: record.linkedIn,
                                },
                            },
                        ],
                        (err, createdRecords) => {
                            if (err) {
                                reject(err);
                            } else {
                                console.log('Record saved to Airtable:', createdRecords[0].getId());
                                logMessage('Record saved to Airtable:' + createdRecords[0].getId());
                                resolve();
                            }
                        }
                    );
                });
            } catch (error) {
                console.error(error);
                throw error; // throwing error if it fails to save record to Airtable
            }
        }
        createOutputDirectory();
        const csvWriter = createCsvWriter({
            path: 'Output/linkedin_profiles.csv',
            header: [
                { id: 'name', title: 'Name' },
                { id: 'description', title: 'Description' },
                { id: 'contact', title: 'Contact' },
                { id: 'picture', title: 'Picture' },
                { id: 'linkedIn', title: 'LinkedIn' },
            ],
        });
        //save data to Output/linkedin_profiles.csv
        await csvWriter.writeRecords(data);
        console.log('Data saved successfully as CSV!');
    } catch (error) {
        console.error('An error occurred during scraping:1', error);
        throw new Error(JSON.stringify(error.message));
    } finally {
        await browser.close();
    }
}

app1.post('/start', async (req, res) => {
    const { slackBotToken, signingSecret, airtableApiKey, airtableBaseId, airtableTableName } = req.body;

    try {
        await saveUsers(slackBotToken, signingSecret, airtableApiKey, airtableBaseId, airtableTableName);
        const airtableLink = `https://airtable.com/${airtableBaseId}`;
        res.send({
            message: 'Scraping process completed successfully.',
            airtableLink: airtableLink,
        });
        logMessage('Scraping process completed successfully.');
    } catch (error) {
        console.error('Final error', error.message);
        res.status(500).send(error.message);
    }
});


io.on('connection', (socket) => {
    console.log('A client has connected.');

    socket.on('disconnect', () => {
        console.log('A client has disconnected.');
    });
});

http.listen(4000, () => {
    console.log('Server is running on port 4000');
});
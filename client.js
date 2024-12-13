// Import libraries
const msal = require('@azure/msal-node');
const axios = require('axios');
const dotenv = require('dotenv');
const readline = require('readline');
const fs = require('fs');
dotenv.config();

// MSAL configuration
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    },
};

const clientApp = new msal.ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
    try {
        const authResult = await clientApp.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
        });
        console.log('Access token acquired successfully.');
        return authResult.accessToken;
    } catch (error) {
        console.error('Error acquiring token:', error);
        throw error;
    }
}

async function createContainer(accessToken, containerName) {
    try {
        const response = await axios.post(
            'https://graph.microsoft.com/v1.0/storage/fileStorage/containers',
            {
                displayName: containerName,
                description: `${containerName} description`,
                containerTypeId: process.env.CONTAINER_TYPE_ID,
            },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                },
            }
        );
        console.log('Container created:', response.data);
        return response.data;
    } catch (error) {
        console.error('Error creating container:', error.response?.data || error);
        throw error;
    }
}

async function listContainers(accessToken) {
    try {
        const response = await axios.get(
            'https://graph.microsoft.com/v1.0/storage/fileStorage/containers',
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                },
                params: {
                    $filter: `containerTypeId eq ${process.env.CONTAINER_TYPE_ID}`,
                },
            }
        );
        console.log('Containers:', response.data.value);
        return response.data.value;
    } catch (error) {
        console.error('Error listing containers:', error.response?.data || error);
        throw error;
    }
}

async function listFilesInContainer(accessToken, containerId) {
    try {
        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/drives/${containerId}/root/children`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                },
            }
        );
        console.log('Files in container:', response.data.value);
        return response.data.value;
    } catch (error) {
        console.error('Error listing files in container:', error.response?.data || error);
        throw error;
    }
}

async function uploadFileToContainer(accessToken, containerId, filePath) {
    try {
        const fileName = filePath.split('/').pop();
        const fileData = fs.readFileSync(filePath);
        const response = await axios.put(
            `https://graph.microsoft.com/v1.0/drives/${containerId}/root:/${fileName}:/content`,
            fileData,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/octet-stream',
                },
            }
        );
        console.log(`File uploaded: ${response.data.name}`);
        console.log(`Copy this URL into your browser to edit: ${response.data.webUrl}`);
        return response.data;
    } catch (error) {
        console.error('Error uploading file:', error.response?.data || error);
        throw error;
    }
}

let activeContainer = null;

function showMenu() {
    console.log("\nSelect an option:");
    console.log("1. Acquire Access Token");
    console.log("2. Create a Container");
    console.log("3. List Containers");
    console.log("4. Set Active Container");
    console.log("5. List Files in Active Container");
    console.log("6. Upload a File to Active Container");
    console.log("7. Exit");
}

async function main() {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    let accessToken = null;

    const prompt = (query) => new Promise((resolve) => rl.question(query, resolve));

    let exit = false;
    while (!exit) {
        showMenu();
        const choice = await prompt("Enter your choice: ");

        try {
            switch (choice) {
                case "1":
                    accessToken = await getAccessToken();
                    break;
                case "2":
                    if (!accessToken) {
                        console.log("Please acquire an access token first (Option 1).\n");
                        break;
                    }
                    const containerName = await prompt("Enter container name: ");
                    await createContainer(accessToken, containerName);
                    break;
                case "3":
                    if (!accessToken) {
                        console.log("Please acquire an access token first (Option 1).\n");
                        break;
                    }
                    const containers = await listContainers(accessToken);
                    console.log('Retrieved containers:', containers);
                    break;
                case "4":
                    if (!accessToken) {
                        console.log("Please acquire an access token first (Option 1).\n");
                        break;
                    }
                    const containersForSetting = await listContainers(accessToken);
                    if (containersForSetting.length === 0) {
                        console.log("No containers available to set as active.\n");
                        break;
                    }
                    console.log("Available Containers:");
                    containersForSetting.forEach((container, index) => {
                        console.log(`${index + 1}. ${container.displayName} (ID: ${container.id})`);
                    });
                    const selectedIndex = await prompt("Enter the number of the container to set as active: ");
                    const selectedContainer = containersForSetting[parseInt(selectedIndex, 10) - 1];
                    if (selectedContainer) {
                        activeContainer = selectedContainer;
                        console.log(`Active container set to: ${activeContainer.displayName}`);
                    } else {
                        console.log("Invalid selection. Please try again.\n");
                    }
                    break;
                case "5":
                    if (!accessToken) {
                        console.log("Please acquire an access token first (Option 1).\n");
                        break;
                    }
                    if (!activeContainer) {
                        console.log("Please set an active container first (Option 4).\n");
                        break;
                    }
                    const files = await listFilesInContainer(accessToken, activeContainer.id);
                    console.log('Files in active container:', files);
                    break;
                case "6":
                    if (!accessToken) {
                        console.log("Please acquire an access token first (Option 1).\n");
                        break;
                    }
                    if (!activeContainer) {
                        console.log("Please set an active container first (Option 4).\n");
                        break;
                    }
                    const filePath = await prompt("Enter the path to the document to upload (default: SimpleSampleDoc.docx): ") || "SimpleSampleDoc.docx";
                    await uploadFileToContainer(accessToken, activeContainer.id, filePath);
                    break;
                case "7":
                    exit = true;
                    console.log("Exiting...");
                    break;
                default:
                    console.log("Invalid choice. Please try again.");
            }
        } catch (error) {
            console.error('Error processing your request:', error);
        }
    }

    rl.close();
}

main();

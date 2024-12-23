// Import libraries
const msal = require('@azure/msal-node');
const axios = require('axios');
const dotenv = require('dotenv');
const readline = require('readline');
const fs = require('fs');
const { showMenu, handleMenuSelection } = require('./client-menu');
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
        // Extract relevant details from the auth result
        const accessToken = authResult.accessToken;
        const expiresOn = new Date(authResult.expiresOn).toLocaleString();
        const scopes = authResult.scopes.join(', ');

        console.log('Access token acquired successfully.');
        console.log(`Token Expires On: ${expiresOn}`);
        console.log(`Scopes: ${scopes}`);
        console.log(`Token (truncated): ${accessToken.substring(0, 20)}...`);
        return authResult.accessToken;
    } catch (error) {
        console.error('Error acquiring token:', error);
        throw error;
    }
}

async function createContainer(accessToken, containerName) {
    const url = 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers';
    const body = {
        displayName: containerName,
        description: `${containerName} description`,
        containerTypeId: process.env.CONTAINER_TYPE_ID,
    };

    try {
        const response = await axios.post(url, body, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
        });
        console.log('Container created:', response.data);
        console.log(`HTTP Method: POST\nURL: ${url}\nBody:`, JSON.stringify(body, null, 2));
        return response.data;
    } catch (error) {
        console.error('Error creating container:', error.response?.data || error);
        throw error;
    }
}

async function grantContainerPermission(accessToken, containerId, email) {
    const url = `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}/permissions`;
    const body = {
        roles: ["owner"],
        grantedToV2: {
            user: {
                userPrincipalName: email,
            },
        },
    };

    try {
        const response = await axios.post(url, body, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
        });
        console.log(`Permission granted successfully to ${email}.`);
        console.log(`HTTP Method: POST\nURL: ${url}\nBody:`, JSON.stringify(body, null, 2));
        return response.data;
    } catch (error) {
        console.error("Error granting permission:", error.response?.data || error);
        throw error;
    }
}

async function listContainers(accessToken) {
    const url = 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers';

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
            params: {
                $filter: `containerTypeId eq ${process.env.CONTAINER_TYPE_ID}`,
            },
        });
        console.log('Containers:', response.data.value);
        console.log(`HTTP Method: GET\nURL: ${url}`);
        return response.data.value;
    } catch (error) {
        console.error('Error listing containers:', error.response?.data || error);
        throw error;
    }
}

async function listFilesInContainer(accessToken, containerId) {
    const url = `https://graph.microsoft.com/v1.0/drives/${containerId}/root/children`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });
        console.log('Files in container:', response.data.value);
        console.log(`HTTP Method: GET\nURL: ${url}`);
        return response.data.value;
    } catch (error) {
        console.error('Error listing files in container:', error.response?.data || error);
        throw error;
    }
}

async function uploadFileToContainer(accessToken, containerId, filePath) {
    const fileName = filePath.split('/').pop();
    const url = `https://graph.microsoft.com/v1.0/drives/${containerId}/root:/${fileName}:/content`;
    const fileData = fs.readFileSync(filePath);

    try {
        const response = await axios.put(url, fileData, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream',
            },
        });
        console.log(`File uploaded: ${response.data.name}`);
        console.log(`Copy this URL into your browser to edit: ${response.data.webUrl}`);
        console.log(`HTTP Method: PUT\nURL: ${url}`);
        return response.data;
    } catch (error) {
        console.error('Error uploading file:', error.response?.data || error);
        throw error;
    }
}

async function inviteUsersToDocument(accessToken, driveId, itemId, recipients, message, requireSignIn = true, sendInvitation = true, roles = ["read"]) {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/invite`;

    const body = {
        recipients: recipients.map(email => ({ email })),
        message: message || "Here's the document we're collaborating on.",
        requireSignIn,
        sendInvitation,
        roles,
    };

    try {
        const response = await axios.post(url, body, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
        });
        console.log(`Invitation sent successfully to ${recipients.join(", ")}.`);
        console.log(`HTTP Method: POST\nURL: ${url}\nBody:`, JSON.stringify(body, null, 2));
        return response.data;
    } catch (error) {
        console.error('Error inviting users to document:', error.response?.data || error);
        throw error;
    }
}

async function showPermissionsOnDocument(accessToken, driveId, itemId) {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        console.log('Permissions on the active document:');
        response.data.value.forEach((permission, index) => {
            console.log(`\nPermission ${index + 1}:`);
            console.log(JSON.stringify(permission, null, 2));
        });
        console.log(`\nHTTP Method: GET\nURL: ${url}`);
        return response.data.value;
    } catch (error) {
        console.error('Error retrieving permissions:', error.response?.data || error);
        throw error;
    }
}

let activeContainer = null;

async function main() {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    let context = {
        accessToken: null,
        activeContainer: null,
        activeDocument: null,
        prompt: (query) => new Promise((resolve) => rl.question(query, resolve)),
        actions: {
            getAccessToken,
            createContainer,
            grantContainerPermission,
            listContainers,
            listFilesInContainer,
            uploadFileToContainer,
            inviteUsersToDocument,
            showPermissionsOnDocument,
        },
    };

    let exit = false;
    while (!exit) {
        showMenu();
        const choice = await context.prompt("Enter your choice: ");
        exit = !(await handleMenuSelection(choice, context));
    }

    rl.close();
}

main();

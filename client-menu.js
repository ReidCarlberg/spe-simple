// Menu and handler for the SharePoint demo application

function showMenu() {
    console.log("\nSelect an option:");
    console.log("1. Acquire Access Token");
    console.log("2. Create a Container");
    console.log("3. List Containers");
    console.log("4. Set Active Container");
    console.log("5. Add User as Owner to Active Container");
    console.log("6. List Files in Active Container");
    console.log("7. Upload a File to Active Container");
    console.log("8. Set Active Document");
    console.log("9. Invite User to Active Document");
    console.log("a. Show Permissions on Active Document");
    console.log("z. Exit");
}

async function handleMenuSelection(choice, context) {
    const { accessToken, activeContainer, prompt, actions } = context;

    switch (choice) {
        case "1":
            context.accessToken = await actions.getAccessToken();
            break;
        case "2":
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            const containerName = await prompt("Enter container name: ");
            await actions.createContainer(accessToken, containerName);
            break;
        case "3":
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            const containers = await actions.listContainers(accessToken);
            console.log('Retrieved containers:', containers);
            break;
        case "4":
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            const containersForSetting = await actions.listContainers(accessToken);
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
                context.activeContainer = selectedContainer;
                console.log(`Active container set to: ${selectedContainer.displayName}`);
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
            const email = await prompt("Enter the email of the user to grant permission: ");
            await actions.grantContainerPermission(accessToken, activeContainer.id, email);
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
            const files = await actions.listFilesInContainer(accessToken, activeContainer.id);
            console.log('Files in active container:', files);
            break;
        case "7":
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            if (!activeContainer) {
                console.log("Please set an active container first (Option 4).\n");
                break;
            }
            const filePath = await prompt("Enter the path to the document to upload (default: SimpleSampleDoc.docx): ") || "SimpleSampleDoc.docx";
            await actions.uploadFileToContainer(accessToken, activeContainer.id, filePath);
            break;
        case "8": // New case
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            if (!activeContainer) {
                console.log("Please set an active container first (Option 4).\n");
                break;
            }
            const filesForSelection = await actions.listFilesInContainer(accessToken, activeContainer.id);
            if (filesForSelection.length === 0) {
                console.log("No files available to set as active.\n");
                break;
            }
            console.log("Available Files:");
            filesForSelection.forEach((file, index) => {
                console.log(`${index + 1}. ${file.name} (ID: ${file.id})`);
            });
            const selectedFileIndex = await prompt("Enter the number of the file to set as active: ");
            const selectedFile = filesForSelection[parseInt(selectedFileIndex, 10) - 1];
            if (selectedFile) {
                context.activeDocument = selectedFile;
                console.log(`Active document set to: ${selectedFile.name}`);
            } else {
                console.log("Invalid selection. Please try again.\n");
            }
            break;
        case "9":
            if (!accessToken) {
                console.log("Please acquire an access token first (Option 1).\n");
                break;
            }
            if (!activeContainer) {
                console.log("Please set an active container first (Option 4).\n");
                break;
            }
            if (!context.activeDocument) {
                console.log("Please set an active document first (Option 8).\n");
                break;
            }
            const recipientEmails = await prompt("Enter recipient emails (comma-separated): ");
            const recipients = recipientEmails.split(',').map(email => email.trim());
            const message = await prompt("Enter an invitation message: ");
            const rolesInput = await prompt("Enter roles (comma-separated, e.g., read,write): ");
            const roles = rolesInput ? rolesInput.split(',').map(role => role.trim()) : ["read"];
            
            const sendInvitationResponse = await prompt("Send invitation email? (yes/no): ");
            const sendInvitation = sendInvitationResponse.toLowerCase() === "yes";
        
            const requireSignInResponse = await prompt("Require sign-in? (yes/no): ");
            const requireSignIn = requireSignInResponse.toLowerCase() === "yes";
        
            await actions.inviteUsersToDocument(
                accessToken,
                activeContainer.id,
                context.activeDocument.id,
                recipients,
                message,
                requireSignIn,
                sendInvitation,
                roles
            );
            break;
        case "a": // New case
        if (!accessToken) {
            console.log("Please acquire an access token first (Option 1).\n");
            break;
        }
        if (!activeContainer) {
            console.log("Please set an active container first (Option 4).\n");
            break;
        }
        if (!context.activeDocument) {
            console.log("Please set an active document first (Option 8).\n");
            break;
        }
    
        const permissions = await actions.showPermissionsOnDocument(
            accessToken,
            activeContainer.id,
            context.activeDocument.id
        );
    
        if (permissions.length === 0) {
            console.log("No permissions found for the active document.\n");
        } else {
            console.log("Permissions:");
            permissions.forEach((permission, index) => {
                console.log(
                    `${index + 1}. Role: ${permission.roles.join(", ")}, User: ${permission.grantedTo?.user?.displayName || "N/A"}`
                );
            });
        }
        break;
    case "z":
            return false;
        default:
            console.log("Invalid choice. Please try again.");
    }
    return true;
}

module.exports = { showMenu, handleMenuSelection };

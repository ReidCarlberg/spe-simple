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
    console.log("8. Exit");
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
        case "8":
            return false;
        default:
            console.log("Invalid choice. Please try again.");
    }
    return true;
}

module.exports = { showMenu, handleMenuSelection };

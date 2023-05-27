# User and Group Creation Script

This project is a self-assignment for school with the objective of creating a script that creates users, creates groups, and adds users to those groups.

## Description

The User and Group Creation Script allows the user to utilize a CSV file to edit the user data of the created users and groups, as well as configure options such as the CSV path, Tenant domain, and Tenant ID.

After each step of the script, multiple checks are performed to ensure smooth progress. In case of any issues, error messages will be displayed to the user.

## Dependencies

Before running the script, make sure you have the following dependencies installed:

- **Powershell:** The script is written in Powershl, so ensure that you have Powersell installed on your system.

 - **AzureAD PowerShell Module**: The script utilizes the AzureAD Powershell module to manage user and group operations in Azure Active Directory. If the module is not already installed, the script will automatically install it for you.

    - For more information about the AzureAD module, visit the [AzureAD PowerShell Gallery page](https://www.powershellgallery.com/packages/AzureAD/).


## CSV Configuration

Before running the script, you need to configure the CSV file with the following values:

-   **First Name**: The first name of the user.
-   **Middle Name**: The middle name of the user.
-   **Surname**: The surname or last name of the user.
-   **Location**: The location of the user.
-   **Department**: The department to which the user belongs.
-   **Function**: The function or role of the user.
-  **Office Location**: The office location of the user.
-   **Security Group**: The security group to which the user should be added.
-   **Password**: The password for the user account.

You can find a template CSV file named "Proftaak - S2 - User & Groups.csv" in the CSV Import folder of this repository. Please use this template as a reference to structure your CSV file accordingly.

## Configuration

Before running the script, you need to configure the following variables:

- **CSV Path**: Specify the path to the CSV file containing users and groups data.

- **Tenant Domain**: Provide the domain name of your Azure AD tenant. This is the domain used to log in to your Azure portal.

- **Tenant ID**: To find the Tenant ID, follow these steps:
  - Log in to the [Azure Portal](https://portal.azure.com/) with your Azure AD admin account.
  - Navigate to **Azure Active Directory**.
  - Click on **Properties** in the left navigation pane.
  - Locate the **Directory ID** field, which represents your Tenant ID.

## Contributors

This project was created as a self-assignment by Boemy. If you would like to contribute to this project, feel free to fork the repository and submit a pull request.

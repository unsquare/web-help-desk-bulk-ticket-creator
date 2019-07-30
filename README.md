# SolarWinds Web Help Desk Bulk Ticket Creator

Although SolarWinds does provide a built-in bulk ticket importer, the implementation is unwieldy if you rely on custom fields.

This Python script is a more user-friendly solution for the bulk ticket creation process. It requires Python 3 to run and a copy of Excel to edit the template.

Before you can start using this script, you will need to create an **API key** for your Web Help Desk tech account. To do so, log in to Web Help Desk and view your account settings, under **Techs > My Account**. You should see a field labeled **API Key**. Click the edit icon and then click **Generate**.

When you run the script for the first time, you will be prompted to enter your email address and API key. The script stores this information in your system keychain.

If you haven't already done so, make sure to install [requests](https://pypi.org/project/requests/), [xlrd](https://pypi.org/project/xlrd/), and [keyring](https://pypi.org/project/keyring/).

You may also need to install [urllib3](https://pypi.org/project/urllib3/) if your Web Help Desk server certificate cannot be verified.

To configure the Excel spreadsheet, create a sheet for each ticket type that you want to support. Make sure to update **definition_ids** and **supported_ticket_types** with the corresponding column names and custom field IDs.

Make sure that the first tab is named **Bulk Ticket Requests**. Use the **Menus** tab to configure data validation for columns on the other sheets. After you finish customizing the template, I recommend protecting the sheets to prevent further edits.
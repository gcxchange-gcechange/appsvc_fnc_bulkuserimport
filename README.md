# appsvc_fnc_bulkuserimport

This function app is part of the bulk user import project.

## Step

1. Funstion app receive a name from a webpart.
2. The Function App Connect to the sharepoint list with the name (Filter list and get id base on title)
3. If find, continue the process, if not find, send an error feedback. Stop there
4. Function app get all list item. 
5. For each item (user):
   1. Update status of list item with: Progress
   2. Check if user exist; If yes, update status with: Error and message: User already exist. Break and start the next item. If not exist, go to next step
   3. Create user invite
   4. Update user with info and usertype
   5. Add user to welcome and assign group
   6. Send invitation
   7. Update status with: complete

During all the process, if an error happen, the status of the item get update with the system error message.

The sharepoint list should be formatted like this:
Field_1: FirstName
Field_2: LastName
Field_3: EmailCloud
Field_4: WorkEmail
Field_5: Status
Field_6: ErrorMessage

All fileld need to be single line expect for ErrorMessage that need to be Multiple line.

## App settings required

* clientId: The application registration create for
* clientSecret: Secret of the application registration
* tenantid: Tenant id where the app is
* BulkSiteId: ID of the site where the sharepoint list is
* welcomeGroup: Id of the welcome group
* gcxAssigned: id og the gcx assign group
* UserSender: Id of the user that send the email

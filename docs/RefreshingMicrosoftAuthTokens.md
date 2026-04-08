
This document serves as a guide to help you recognize when you need to refresh your authentication token, and how to do it.

## Symptom

Occasionally, you might see an error message like this when authenticating

> The credentials provided are invalid. Reason: AADSTS700082: The refresh token has expired due to inactivity. The token was issued on {issueDate} and was inactive for 90.00:00:00.
> https://login.microsoftonline.com/error?code=700082


![47900805-1403-4425-9ef4-d8638833c499](https://github.com/user-attachments/assets/5d40025c-28fd-4b0b-b138-ece003138fc4)


This just means you've gone 90 days without using the token it has cached.  This is expected whenever you use a DaqBook laptop that you last used over 90 days ago.  As Microsoft states, this is an

> Expected part of the token lifecycle - the user went an extended period of time without using the application, so the token was expired when the app attempted to refresh it.

## Resolution

To resolve this, you'll need to clear out the old tokens.  Once the old tokens are cleared, Excel will prompt you to reauthenticate next time the queries refresh.

### Steps

1. With your TUS workbook open in Excel, go to `Data > Get Data > Data Source Settings...`

<img width="428" height="629" alt="image" src="https://github.com/user-attachments/assets/de530887-3745-4566-9b63-e999f014109a" />


2. In the **Data source settings** window, click `Clear All Permissions` then `Delete`, then `Close` the **Data source settings** window.


<img width="802" height="584" alt="image" src="https://github.com/user-attachments/assets/4c462d82-4683-4c49-8b73-05f6aa224400" />

<img width="462" height="184" alt="image" src="https://github.com/user-attachments/assets/398d35c9-8e1e-4d9c-b4c0-e4041c5e10a2" />


3. Now, go to `Data > Refresh All`.  Excel will automatically prompt you to reauthenticate.


<img width="570" height="167" alt="image" src="https://github.com/user-attachments/assets/5409b9c6-12b6-476b-acdc-20c7dec00935" />


4.  Follow the usual steps for authentication.

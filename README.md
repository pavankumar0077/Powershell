## Script Flow ( Architecture )
``` mermaid
flowchart TD
    A[User] -->|Enter API URL & Method| B[Script Start]
    B --> C{Method = POST or PUT?}
    C -->|Yes| D[Payload Input: Direct JSON or File Picker]
    C -->|No| E[No Payload Needed]
    D --> F
    E --> F
    F[Invoke-RestMethod - API Call] -->|Response JSON| G[Flatten JSON]
    G --> H[Clean Column Names remove :data]
    H --> I[Add Timestamp]
    I --> J{Is Excel File Open?}
    J -->|Yes| K[Save to New Temp File API_Response_YYYYMMDD.xlsx]
    J -->|No| L[Save/Append to API_Response.xlsx]
    K --> M[Export to Excel]
    L --> M[Export to Excel]
    M --> N[Results Saved]
    N --> O{Open Excel Now?}
    O -->|Yes| P[Launch Excel File]
    O -->|No| Q[End]
```

## POST 

<img width="1337" height="461" alt="image" src="https://github.com/user-attachments/assets/f13fbd73-1c70-47c9-bb9f-0ca0c2563827" />

<img width="1380" height="784" alt="image" src="https://github.com/user-attachments/assets/45548b62-658e-447d-b107-39026b8aa0b7" />


## GET

<img width="1561" height="798" alt="image" src="https://github.com/user-attachments/assets/6c117958-10a0-4f43-9228-e9742a4d05ec" />


## TALEND API TESTER - EXTENSION
``` https://chromewebstore.google.com/detail/talend-api-tester-free-ed/aejoelaoggembcahagimdiliamlcdmfm ```


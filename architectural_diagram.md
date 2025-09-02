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
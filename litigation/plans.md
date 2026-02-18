PLANS FOR TWEAK OR UPDATES:
- adding a api key modal for user encode
- adding a gemini model view for them to select what AI model they will be using where we need to also add a highlight to the latest one needed for the system.
    - use the https://generativelanguage.googleapis.com/v1beta/models?key=YOUR_API_KEY for fetching list of AI model lists for gemini
    - will be used so that any gmeini model the user will pick the api will still be working. 
- making sure that the system or the appscript will continuously work without editing always the code. another 
- A log sheet to display any errors that the system will encounter so that the user that will run the appscript will use that report for reporting to the developer.

- Testing should diagnose the systems integration with gemini ai model and if it really is working. 

``` 
REST

curl "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent" \
  -H "x-goog-api-key: $GEMINI_API_KEY" \
  -H 'Content-Type: application/json' \
  -X POST \
  -d '{
    "contents": [
      {
        "parts": [
          {
            "text": "Explain how AI works in a few words"
          }
        ]
      }
    ]
  }'

```


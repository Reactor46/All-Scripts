## Angular2 Windows Servers status
> This project is generated with [generator-angular2-typescript](https://github.com/shibbir/generator-angular2-typescript) version 0.10.6.

The PowerShell script is gathering information from AD for all MS Windows Servers and saves it to ping.json and servers.json

## Installation

- Clone the repository.
- Install npm dependencies.
- Build the Angular2 app (See: Production Build bellow)
- When built copy "dist" folder to webserver.
- Copy info.bat and info.ps1 into "dist" folder.
- Start PS script "info.ps1" or "info.bat" BATCH file to start information gathering. Once done you should have both json files generated in "assets" sub-folder

```Windows CLI (Command prompt)
C:\NG2ServerStatus\npm install -g typescript
C:\NG2ServerStatus\npm start
```
## Production Build
```Windows CLI (Command prompt)
C:\NG2ServerStatus\npm run build
```
## Running Unit Tests
```Windows CLI (Command prompt)
C:\NG2ServerStatus\npm test
```

## Running End-to-End Tests
```Windows CLI (Command prompt)
REM make sure you have a running app
C:\NG2ServerStatus\npm run e2e
```

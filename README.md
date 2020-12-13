<!-- PROJECT LOGO -->
<p align="center">
  <a href="https://github.com/ckzgraphics/ScriptUtlity">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Script Utility</h3>

  <p align="center">
    Download the data!!!
  </p>
</p>

#### Prerequisite
- Java 8
- Chrome Browser
- Good internet connection

#### Setup
- Download the project
- Keep the excel file in "data" directory

#### Windows Execution
- Open command prompt and navigate to the home directory of the project. For eg- 
    ```sh
    cd "D:/Projects/setup/ScriptUtlity"
    ```
- Execute the following command:
   ```sh
   ./mvnw.cmd clean compile -Dbook=<file-with-extn> -Durl="<website-url>"
   ```

#### Mac Execution
- Open terminal and navigate to the home directory of the project. For eg- 
    ```sh
    cd Documents/Projects/setup/ScriptUtlity
    ```
- Execute the following command:
   ```sh
   ./mvnw clean compile -Dbook=<file-with-extn> -Durl="<website-url>"
   ```

##### Example
- Consider file name is `POC.xlsx` and website url is `https://www.abc.com/symbol=`
- The execution command for mac would be-
    ```sh
    ./mvnw clean compile -Dbook=POC.xlsx -Durl="https://www.abc.com?symbol="
    ```
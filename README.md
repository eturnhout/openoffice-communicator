# openoffice-communicator
Mostly made this to convert word documents to html (I may add more in the future). It needs a lot of improvements, but I'm happy for now.

## Requirements
* Maven
* OpenOffice/LibreOffice running in headless mode e.g. ``soffice.exe -headless -accept="socket,host=0.0.0.0,port=8100;urp;" -nofirststartwizard`` on windows

## Testing
To start testing create a file called ``config.properties`` in the root directory of the project and add the following configurations:

  ```
  host=OpenOfficeHostMachine
  port=PortToOpenOffice
  folder=ExistingFolder
  ```

Now execute ``mvn clean test`` and the test should pass.

## Usage
Build it with ``mvn install`` and you can use it as a dependency for another project. Add the following to your pom.xml file's dependencies to start using it:

  ```
  <dependency>
    <groupId>dev.evt.openoffice</groupId>
    <artifactId>openoffice-communicator</artifactId>
    <version>1.0-SNAPSHOT</version>
  </dependency>
  ```
  
  See the DocumentManagerTest#testTextDocument to see how it's used. (Will be documented properly when more features are addded)

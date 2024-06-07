# Word to PDF Converter

This project is a Java-based desktop application that converts Microsoft Word documents (.doc and .docx) to PDF files. It uses the docx4j library for handling Word documents and Apache FOP for PDF conversion. The graphical user interface (GUI) is built using Swing and styled with FlatLaf.

## Features

-   Add multiple Word documents for batch conversion.
-   Select an output directory for the converted PDF files.
-   Supports both .doc and .docx formats.
-   Maps fonts to ensure proper rendering in the PDF files.

## Prerequisites

-   Java 8 or higher
-   Maven 3.6 or higher

## Getting Started

### Clone the Repository

```bash
git clone https://github.com/Hikaru-e/docx2pdf.git
```

```bash
cd docx2pdf 
```
### Build the Project

Use Maven to build the project and create a runnable JAR file.

```bash
mvn clean package
```
### Run the Application

After building the project, you can run the application using the generated JAR file located in the `target` directory.

```bash
java -jar target/docx2pdf-1.0-SNAPSHOT-shaded.jar 
```

![Running Application](https://github.com/Hikaru-e/docx2pdf/assets/77628961/2ff9a562-0a09-4a95-9dd7-5bfff6e02c4b)

## Usage

1.  Launch the application.
2.  Click on "Add Files" to select Word documents (.doc or .docx) for conversion.
3.  Click on "Browse" to choose an output directory where the converted PDF files will be saved.
4.  Click on "Convert to PDF" to start the conversion process.
5.  The application will notify you upon successful conversion of each file.

## Project Structure

-   `src/main/java/org/example/Main.java`: The main class containing the application logic and GUI.
-   `src/main/resources/`: Contains additional resources such as icons and styles.
-   `pom.xml`: Maven project file containing dependencies and build configuration.

## Dependencies

The project relies on the following main dependencies:

-   **docx4j**: For handling Word documents.
-   **Apache FOP**: For converting documents to PDF.
-   **FlatLaf**: For modern look and feel of the GUI.

All dependencies are defined in the `pom.xml` file and are managed by Maven.


## Acknowledgements

-   [docx4j](https://github.com/plutext/docx4j)
-   [Apache FOP](https://xmlgraphics.apache.org/fop/)
-   [FlatLaf](https://www.formdev.com/flatlaf/)
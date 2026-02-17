Excel to Word Converter (Java + Apache POI)

 A simple Java application that converts Excel (.xlsx) data into a formatted Word (.docx) table using Apache POI.
 This project demonstrates file handling, Apache POI usage, and document generation in Java.
 
Features

Read data from Excel (.xlsx)

 Convert Excel rows into Word table format

 Preserve row and column structure

Handle different cell types:

 String

 Numeric

 Boolean

 Formula

 Date

 Automatically generate Word document

 Tech Stack

 Java

 Apache POI

 IntelliJ IDEA

 Maven

Project Structure
 Excel-To-Word-Converter
 │
 ├── src
 │   └── main/java/org/example/Main.java
 │
 ├── data
 │   └── Employee details.xlsx
 │
 ├── output
 │   └── Employee list.docx (Generated)
 │
 └── pom.xml

Maven Dependencies

Add this inside your pom.xml:

<dependencies>

    <!-- Excel (.xlsx) -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.5</version>
    </dependency>

    <!-- Word (.docx) -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>5.2.5</version>
    </dependency>

</dependencies>

How to Run

Clone the repository

Open in IntelliJ IDEA

Place your Excel file inside the data folder

Update file path if needed in Main.java:

String excelFilePath = "data/Employee details.xlsx";
String wordFilePath = "output/Employee list.docx";


Run Main.java

The Word file will be generated inside the output folder.

Example Excel Input
Emp ID	Name	Department	Salary
101	Aisha Khan	HR	45000
102	Rahul Verma	IT	65000

Future Improvements

Add header styling (bold, background color)

Auto column width adjustment

Convert multiple sheets

Build Spring Boot REST API version

Add file upload feature

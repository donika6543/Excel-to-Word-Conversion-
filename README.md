# Excel to Word Converter (Apache POI)

This project is a Java-based application that converts data from an **Excel (.xlsx)** file into a **Word (.docx)** document using the Apache POI library.

It demonstrates how to read Excel files and generate formatted Word documents programmatically using Java.

---

## 🚀 Features

- Read data from Excel (.xlsx) files
- Convert Excel rows into formatted Word (.docx) content
- Automatically create tables in Word from Excel data
- Simple and clean Java implementation
- Useful for reports and document automation

---

## 🛠️ Technologies Used

- Java
- Apache POI
- Maven
- IntelliJ IDEA (recommended)

---

## 📦 Maven Dependencies

Add the following dependencies to your `pom.xml`:

```xml
<dependencies>

    <!-- Excel (.xlsx) support -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.5</version>
    </dependency>

    <!-- Word (.docx) support -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>5.2.5</version>
    </dependency>
</dependencies>





keeps-characterization-msoffice
===============================

Characterization tool of Microsoft Office files (doc, docx, ppt, pptx, xls, xlsx) using Apache POI

## Build & Use

To build the application simply clone the project and execute the following Maven command: `mvn clean package`  
The binary will appear at `target/msoffice-characterization-tool-1.0-SNAPSHOT-jar-with-dependencies.jar`

To see the usage options execute the command:

```bash
$ java -jar target/msoffice-characterization-tool-1.0-SNAPSHOT-jar-with-dependencies.jar -h
usage: java -jar [jarFile]
 -f <arg> file to analyze
 -h       print this message
 -v       print this tool version
```

## Tool Output Example
```bash
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ooxmlValidationResult>
    <valid>true</valid>
    <word>
        <creator>Sebastien Leroux</creator>
        <description>Test comments</description>
        <keywords>Test keywords</keywords>
        <revision>1</revision>
        <subject>Test subject</subject>
        <title>Test title</title>
        <lineCount>0</lineCount>
        <parCount>0</parCount>
        <sectionCount>0</sectionCount>
        <editTime>0</editTime>
        <osVersion>0</osVersion>
        <wordCount>0</wordCount>
        <pageCount>0</pageCount>
    </word>
</ooxmlValidationResult>
```

## License

This software is available under the [LGPL version 3 license](LICENSE). All corpora was created specifically for this project and is available under the [Creative Commons Attribution-ShareAlike 3.0 Unported License](http://creativecommons.org/licenses/by-sa/3.0/deed.en_US").




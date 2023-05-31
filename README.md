# Moodle Rubric Importer
![Moodle Rubric Importer logo: A clipboard icon filled with a green gradient](images/icon-128.png)

This unofficial extension adds an "Import Excel" button when editing a Moodle rubric.

The expected Excel format should have two rows per criterion. The first column of a criterion should contain its name / description. After that, each column should describe a level for that criterion, with the first row containing its definition and the second its score.

A number of rows at the beginning can be offset (skipped) to account for rubric headers. For each criterion, levels are dynamically included until an empty column is found. Criteria are added until an empty row is found.

An example of the expected rubric format can be seen in the screenshots.

Note that the extension simulates the process of defining a rubric by programatically clicking the buttons, and has only been tested on a specific deployment of Moodle. It may take a couple of seconds to process a rubric, especially with larger files.

If you encounter any issue please contact the developer.

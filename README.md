**Disclaimer**: This is a personal project with no intention of commercial gain.

# G-Suite Integrated Image Snippet Generator
This project aims to generate snippets of images from PDF or Image files from the comfort of one's own Google Drive. The application is designed to use the commenting featue of Google Drive to extract specific areas of images from PDFs or Images. This is a complete cloud-based project built primarily on Google Suite of applications with the additon of Microsoft Power Automate as the main processing unit. 

## Tools and Technologies: 
- Google App Scripts, Microsoft Power Automate
## Services: 
- Encodian Image Extractor from PDF, PDF merge/split code from [pokyCoder](https://stackoverflow.com/users/11904337/pokycoder).
- G-Suite of applications--Drive, Sheets, Slides

## Instructions - [Demo](https://youtu.be/nqtWYWODFOM)
1. Upload files to google drive
2. Set there share status to public
3. Select required sub-snippet regions on PDF or Image files with the comment's adhering to the [syntax mentioned](https://docs.google.com/document/d/1pbxdFL0Z2f1iqIg02PC6GQNMbWNAZ3l2PG5UnI94kio/edit).
4. Copy the required file's shareable link and paste it to the [interfacing form](https://docs.google.com/forms/d/e/1FAIpQLSdeTr5naEzs5xaJcmGYqy_Dqyl2wRxsvwi4VldxG1ReuS9Z4Q/viewform?usp=sf_link).
5. Specify the mail you'd want to recieve real-time updates of the backend.
6. Enjoy image snippets [here](https://drive.google.com/drive/u/0/folders/1uaxct9hfVcGow2e_vi92_jYOdfqv0Nr0) :) (This output folder is subjected to change in future)

# System Design

## Functioning
![Sequence Diagram for Functioning](https://github.com/rohan-motukuri/image-snippets-generator/assets/123802857/641e6761-bf33-4c43-a7e5-1abb7a0b1e08)

# Syntax
The syntax depends on a custom note_taking-cum-documentation syntax I came up with during my first and second year at college for the purpose of digitalizing my physical notes. You can find more insights into the same from [this document](https://docs.google.com/document/d/1pbxdFL0Z2f1iqIg02PC6GQNMbWNAZ3l2PG5UnI94kio/edit).


# Achievements
- Created [full fledged notes](https://docs.google.com/document/d/1hzhBF_XRHfZIbZonC-1VUiR-4oxPZuU-4PDnxhwsgjw/edit?usp=sharing) with an aid from this project during my 2nd year of B.Tech.

# Road map for Improvements
- As this was my first ever project the code quality isn't in the most readable and managable format. Will update it soon with my latest experience and knowledge.
- Currently lacks documentation of proper usage syntax and technical details. Will work on after refractoring the code.
- Bug: It's throwing error for single page PDFs and multiple files.
- Usage: The actual snippet generation can take a long time (approximately 40 Mins if not in test mode) due to the non-standard approach at making a Web Hook from Calendar API and Calendar Connector.
- Standardize the tech stack

  # Attributions
  - PDF Merge and Split code by [pokyCoder](https://stackoverflow.com/users/11904337/pokycoder).
  - Free (for professional use) [Encodian](https://www.encodian.com/products/flowr/) Images From PDF Power Automate Connector.

This project is a Preventative Dropout Dashboard that allows users to upload student data in Excel or CSV format and analyze it to identify students who may be at risk of dropping out. The goal is to turn raw school data such as attendance, grades, and engagement into clear, actionable insights that can support early intervention.

The application is designed to be simple and accessible. Users can upload one or more files, view and manage those uploads, and explore structured results generated directly in the browser. The interface organizes information using a tabbed layout so different datasets and views remain easy to navigate. The system also supports adding notes or comments for individual students, making it possible to track observations and follow-up actions in a more organized way.

All data processing currently happens on the client side, which means users can open the application in their browser and use it immediately without installing anything or creating an account. Each session is independent, so uploaded data and notes are not permanently stored and may need to be reloaded if the page is refreshed. This design keeps the application fast and easy to use while still demonstrating the full workflow of data upload, analysis, and insight generation.

The project is built using React, Vite, and TypeScript, with SheetJS used for parsing Excel files. It is deployed as a web application and can be accessed through a public link, allowing anyone to use it directly from their browser.

To run the project locally, clone the repository, install dependencies using npm install, and start the development server with npm run dev. To create a production build, use npm run build and deploy the output using a hosting platform.

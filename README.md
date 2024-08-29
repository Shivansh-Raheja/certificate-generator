# Certificate Generation and Emailing Server

This Node.js server is designed to generate participation certificates using Google Slides, export them as PDFs, and send them via email to participants listed in a Google Sheets document. The server also provides logging functionality to track the number of certificates generated.

## Features

- Generate certificates from a Google Slides template.
- Export certificates as PDF files.
- Send certificates via email with a customizable email body.
- Fetch logs of generated certificates.

## Prerequisites

Before you begin, ensure you have met the following requirements:

- Node.js installed (version 14.x or later).
- Google API credentials with access to Google Sheets, Google Drive, and Google Slides.
- A Gmail account for sending emails via Nodemailer.
- A Google Slides template for generating certificates.
- `.env` file containing necessary environment variables.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/your-repository.git
    ```

2. Navigate to the project directory:

    ```bash
    cd your-repository
    ```

3. Install dependencies:

    ```bash
    npm install
    ```

4. Create a `.env` file in the root directory and add the following variables:

    ```plaintext
    PORT=3000
    CLIENT_ID=your-client-id
    CLIENT_SECRET=your-client-secret
    REDIRECT_URI=your-redirect-uri
    REFRESH_TOKEN=your-refresh-token
    TEMPLATE_ID=your-google-slides-template-id
    FOLDER_ID=your-google-drive-folder-id
    ```

## Usage

### Starting the Server

To start the server, run:

```bash
npm start

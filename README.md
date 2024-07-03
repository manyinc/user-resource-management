# TAU Data Base Management Application

## Overview

This Python application provides a user interface for managing a database using various utilities such as generating QR codes, creating new users, generating protocols, and more. It utilizes custom widgets from `customtkinter` for an enhanced UI experience.

## Installation

### Prerequisites

- Python 3.x installed
- Dependencies listed in `requirements.txt`

### Installation Steps

1. Clone the repository:
```bath
git clone https://github.com/manyinc/user-resource-management.git
cd user-resource-management
```

2. Install dependencies:
```bath
pip install -r requirements.txt
```

4. Run the application:
```bath
python your_app_file.py
```

## Features

- **New User**: Allows creating a new user and adding them to the database.
- **QR Code**: Generates QR codes with customizable options.
- **Addons**: Provides options to generate various add-ons like footers, `.vcf` files, send emails, print data, and more.
- **User Info**: Displays detailed information about users stored in the database.
- **Protocol**: Generates different types of protocols based on user actions.

## Usage

1. Launch the application.
2. Use the buttons in the left panel to navigate and perform actions:
- **New User**: Fill in user details and click "Add user to database".
- **QR Code**: Enter link and name, select color option, and click "Generate QR code".
- **Addons**: Check desired options (Generate:mail signature, contact file, send them by mail to selected user, print login and password, send bookmarks) and click "Complete" to generate.
- **User Info**: Search for users by name or data and view details.
- **Protocol**: Enter user ID and select protocol type, then click "Generate Protocol".

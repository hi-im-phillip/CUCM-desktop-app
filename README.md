# CUCM Functions

This Python script provides a graphical user interface (GUI) for performing various operations related to Cisco Unified Communications Manager (CUCM). The script utilizes the Tkinter library for the GUI components.

## Requirements

- Python (version 3.x)
- Tkinter library
- Requests library
- Suds library
- Pyperclip library
- PyAutoGUI library

Ensure you have these libraries installed before running the script.

## Usage

1. Run the script using a Python interpreter.
2. Enter the required information in the provided text fields:
   - **Ime i prezime:** Agent's first and last name.
   - **Korisničko ime:** Agent's username.
3. Choose the desired action using the provided buttons:
   - **Odjaviti isk agenta:** Deactivate ISK agent.
   - **Odjaviti ph agenta:** Deactivate PH agent.
   - **Slobodne ekstenzije isk:** Display available ISK extensions.
   - **Slobodne ekstenzije ph:** Display available PH extensions.
   - **Zatvoriti:** Close the application.

## Functions

### ISK Agent Deactivation
- **Function:** `deactivation_agent_isk_GUI`
- **Parameters:**
  - `first_last_name` (string): Agent's first and last name.
  - `username` (string): Agent's username.

### PH Agent Deactivation
- **Function:** `deactivation_agent_ph_GUI`
- **Parameters:**
  - `first_last_name` (string): Agent's first and last name.
  - `username` (string): Agent's username.

### ISK Extensions Retrieval
- **Function:** `show_isk_extensions`
- **Action:** Display available ISK extensions in the text box.

### PH Extensions Retrieval
- **Function:** `show_ph_extensions`
- **Action:** Display available PH extensions in the text box.

## Note

- The script interacts with the CUCM using AXL (Administrative XML) API.
- Connection details and credentials are provided in the script.
- Ensure a VPN connection is established for proper functionality.
- Copy functionality is enabled for the displayed information in the text box.

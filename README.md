# Task 2 – Word JS Add-in (Search Interface)

## Overview

This project implements a **single-page Word JavaScript Add-in** that provides search capabilities inside a Word document.  
The Add-in runs as a **Task Pane** within Microsoft Word and is built using **React 18**, **TypeScript**, and **Chakra UI**.

The goal of this task is to demonstrate how a React-based UI can interact with a Word document using the **Word JavaScript API**.

---

## Features

- Text input for entering a search query
- Case-sensitive search toggle (on / off)
- Executes search using `Word.run` and `document.body.search`
- Displays the **top 3 search results** inside the Add-in panel
- Clears the input field after each search
- Handles loading and empty states

---

## Tech Stack

- React 18
- TypeScript
- Chakra UI
- Office.js / Word JavaScript API
- Microsoft Word Add-in (Task Pane)

---

## Assumptions & Design Decisions

- The implementation focuses only on the search interface, as required.
- Only the top 3 search results are displayed in the Add-in, even if more matches exist.
- Search results are presented as text snippets for simplicity.
- The search input is cleared after each search to improve usability.
- No document modifications (such as highlighting or navigation) were implemented, as they were not part of the requirements.

---

## Challenges & How They Were Overcome

- **Debugging inside a Word Add-in**  
  Standard browser DevTools are limited inside Word.  
  This was handled using terminal logs and temporary UI-based debugging.

- **Office.js lifecycle**  
  Word APIs are only available after `Office.onReady()` is called.  
  The React application is rendered only after Office is fully initialized.

- **Learning and understanding new documentation**  
  The official Microsoft documentation was used extensively to understand the Word JavaScript API and Add-in lifecycle.

---

## Prerequisites

- Node.js (v16+ recommended)
- Microsoft Word (Desktop – Mac or Windows)

---

## Installation & Run

Install dependencies:

```bash
npm install

## Instaltion

1. Install the dependencies:

   `npm ci`

2. Run the development server:

   `npm run dev-server`

3. Side load the manifest.xml file into Word (Desktop)
   `npm run start:desktop`

4. Or Side load the manifest.xml file into Word (Web)
   `npm run start:web`

5. Stop the Word Side Loading
   `npm run stop`
```

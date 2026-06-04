const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign, AlignmentType, UnorderedList, LevelFormat } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Arial", size: 22 }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "264653" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2f5a6b" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 }
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "5a6e76" },
        paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          },
          {
            level: 1,
            format: LevelFormat.BULLET,
            text: "◦",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("40 Thrive Challenge App")]
      }),
      new Paragraph({
        children: [new TextRun("Complete Project Inventory & Technical Documentation")]
      }),
      new Paragraph({
        children: [new TextRun("")]
      }),
      new Paragraph({
        children: [new TextRun("Generated: June 4, 2026")]
      }),
      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== PROJECT OVERVIEW ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("1. PROJECT OVERVIEW")]
      }),
      new Paragraph({
        children: [new TextRun("The 40 Thrive Challenge is a Firebase-based habit tracking application built as a single-page HTML/JavaScript application. Users sign up, select 3 daily habits (one from each category: Sleep, Nutrition, Movement), and track them daily for 40 consecutive days. The app includes auto-reset logic for missed days with an exception for strength training workouts.")]
      }),
      new Paragraph({
        children: [new TextRun("")]
      }),
      new Paragraph({
        children: [new TextRun("Current Status: Two versions exist:")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("25 KB version (outputs/index.html) - Lean rebuild with core features")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("60 KB version (uploads/index.html from 9:23) - Feature-rich version with UI polish")]
      }),
      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== PAGES/SCREENS ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("2. PAGES & SCREENS")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Current 25 KB Version")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Screen-Hero (screen-hero): Hero page with login/signup forms, May Tom bio card with photo, YouTube link, Google Calendar booking")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Status: PRESENT")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Missing from full version: No scrollable content over background photo, no challenge description box, no site-bar at top")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Screen-Onboarding (screen-onboarding): Habit selection - user picks ONE habit from each of three categories")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Status: PRESENT")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Present in both versions")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Screen-App (screen-app): Main application screen with daily tracker, settings, and navigation tabs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Status: PRESENT")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Present in both versions but with different tab structures")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Old 60 KB Version (PRESERVES)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Screen-Loading (screen-loading): Initial loading state shown before auth state is resolved")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Status: MISSING from 25 KB version")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should be restored")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Screen-Auth (screen-auth): Full-screen hero section with scrollable content over background photo")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Status: SUPERSEDED by Screen-Hero in 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Elements to preserve: Challenge description card with features, scrollable content with fixed background, more detailed May Tom bio")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== COMPONENTS ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("3. UI COMPONENTS")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Navigation & Layout")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Site Bar (top navigation) - MISSING FROM 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: Teal background, logo with brand name, navigation links")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should be restored at top of all screens")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bottom Navigation Tabs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("25 KB: 'Today', 'Settings', 'Back' buttons")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: 'Today' (☀️), 'Progress' (📊) tabs with nav icons")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Progress Tab - MISSING FROM 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: Shows 40-day progress grid with visual cells, week-view tracker for strength training")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should be restored")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Auth Modal - MISSING FROM 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("60 KB uses modal-based auth that slides in from right side of screen")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Modal ID: modal-auth")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Tabs: 'Join the Challenge' (Signup) and 'Sign In'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Magic Link Option: 'Send me a sign-in link instead' button with email link sign-in")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("25 KB uses inline forms on hero page - less polished UX")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should be upgraded to modal system")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Bio/Challenge Cards")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("May Tom Bio Card - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Photo: https://lh3.googleusercontent.com/d/1aDp3lAkYmbOGLzLeRcBxysJYLsTYuS5k")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Calendar Link: https://calendar.app.google/tThrrV9Lk9fuqdKv7")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("YouTube Link: https://www.youtube.com/@maytomrd")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Challenge Description Card - MISSING FROM 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: Semi-transparent card with 40-day challenge explanation, three habit categories with emoji and descriptions")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should be restored")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Daily Tracker")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Habit Checkboxes - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Three checkbox toggles for Sleep, Nutrition, Movement habits")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Progress Bar - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Shows percentage of habits completed for today")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Day Badge/Streak Display - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Shows current day (1-40) and streak counter")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Strength Training Week View - MISSING FROM 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: Seven day-of-week dots showing strength training completion for week")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Needed for auto-reset exception logic")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Settings Panel")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Email List Section - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Google Form integration for newsletter signup")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Reset Challenge Button - PRESENT in both")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Resets to Day 1 with confirmation modal")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Sign Out Button - PRESENT in both")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== DATABASE STRUCTURE ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("4. DATABASE STRUCTURE")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Firebase Project")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Project ID: thrive-850e5")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Auth Domain: thrive-850e5.firebaseapp.com")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Database: Firestore")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Collection: users")]
      }),
      new Paragraph({
        children: [new TextRun("Document ID: User UID (from Firebase Authentication)")]
      }),
      new Paragraph({
        children: [new TextRun("")]
      }),
      new Paragraph({
        children: [new TextRun("Fields:")]
      }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                shading: { fill: "264653", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Field", bold: true, color: "FFFFFF" })] })]
              }),
              new TableCell({
                shading: { fill: "264653", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Type", bold: true, color: "FFFFFF" })] })]
              }),
              new TableCell({
                shading: { fill: "264653", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Purpose", bold: true, color: "FFFFFF" })] })]
              }),
              new TableCell({
                shading: { fill: "264653", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Example", bold: true, color: "FFFFFF" })] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("email")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("String")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("User's login email")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("user@example.com")] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("createdAt")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Timestamp")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("When user signed up (Day 1)")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("2024-06-04T...")] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("habits")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Object")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("User's selected habits (indices 0-2)")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("{sleep:0, nutrition:1, movement:2}")] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("startDay")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Number")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Current day (1-40). Resets to 1 on missed day")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("5")] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("dailyLogs")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Object<Map>")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("Daily habit completion status by date")] })]
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun("See below")] })]
              })
            ]
          })
        ]
      }),

      new Paragraph({ children: [new TextRun("")] }),

      new Paragraph({
        children: [new TextRun("dailyLogs Structure (keyed by ISO date string YYYY-MM-DD):")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'2024-06-04': { sleep: false, nutrition: true, movement: true }")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Each date contains boolean for each habit category")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("60 KB Version Additions")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("startDate (instead of createdAt): ISO date string")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("logs (instead of dailyLogs): Same structure")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("strengthSessions: Object tracking dates of 3x/week strength training")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("completed: Boolean flag (not used yet)")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== AUTHENTICATION FLOW ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("5. AUTHENTICATION FLOW")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Firebase Setup")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Auth Method: Email/Password (both versions)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Password Reset: sendPasswordResetEmail (both versions)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Magic Link (60 KB ONLY): sendSignInLinkToEmail + signInWithEmailLink")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Sign-Up Flow")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("User enters: Full Name, Email, Password, opt-in checkbox")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Firebase: createUserWithEmailAndPassword()")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Firestore: Create user doc with email, createdAt timestamp")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Google Form: submitToGoogleForm() sends name, email, opt-in for newsletter")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Redirect: Show 'Welcome' modal → Onboarding screen")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Sign-In Flow")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Standard: signInWithEmailAndPassword()")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Magic Link (60 KB): sendSignInLinkToEmail() → User clicks email link")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Load user state from Firestore")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Check for missed days → Auto-reset if needed")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Route to app screen")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Auth State Listener")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("onAuthStateChanged() fires on page load")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("If logged in: Load user data from Firestore → Call handleAuthState()")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("If not logged in: Show hero/auth screen")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== FEATURES - CURRENT ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("6. CURRENT FEATURES")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("✓ Implemented in 25 KB Version")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("User Authentication: Email/password signup and login")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Habit Selection: Choose 1 from 3 categories (Sleep, Nutrition, Movement)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Daily Tracker: Checkbox toggles for 3 habits")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Progress Bar: Shows % of today's habits completed")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Auto-Reset Logic: Resets to Day 1 if any daily habit not completed")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("EXCEPTION: If strength training (movement=2), allows unchecked days if 3+ sessions completed that week")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bio Section: May Tom's photo, title, bio text, YouTube link, Google Calendar booking button")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Google Form Integration: Email capture for newsletter signup")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Settings Panel: Reset challenge button, sign-out button")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Responsive Design: Mobile-first approach")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Day Counter: Shows current day (1-40) with streak display")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("✓ Implemented in 60 KB Version (Additions)")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Magic Link Sign-In: Email link authentication option")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Modal Auth System: Polished slide-in modal with signup/login tabs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Challenge Description: Detailed card explaining 40-day challenge with category descriptions")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Fixed Hero Background: Scrollable content over background photo")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Progress Tab: 40-day visual grid showing completion status")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Strength Training Week Tracker: Visual dots showing which days strength sessions completed")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Site Bar: Top navigation with branding")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("updateDoc Support: Partial Firestore updates (more efficient than setDoc)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Status Cards: Stats display (current day, completed days, longest streak)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Success Modals: Animated 'all set' and 'all done today' feedback modals")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Enhanced Typography: Better spacing, emoji icons, more polished styling")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== HABIT DEFINITIONS ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("7. HABIT DEFINITIONS")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Sleep & Recovery (🌙)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 0: Go to bed & wake at the SAME time every day (7-8 hours)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 1: 10 min morning sunlight within 60 min of waking")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 2: No screens 1 hour before bed")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Nutrition (🌿)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 0: Drink 20 oz of filtered water upon waking")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 1: Eat 5 DIFFERENT vegetables per day")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 2: 30+ grams of protein at breakfast")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Movement (💪)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 0: 40 squats & 20 push-ups after each meal")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 1: 7,500 steps in a day")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Option 2: 3 strength training sessions/week (30 min min) ← Special auto-reset exception")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== AUTO-RESET LOGIC ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("8. AUTO-RESET LOGIC")]
      }),

      new Paragraph({
        children: [new TextRun("The app automatically resets users to Day 1 if they miss a day. However, if users select strength training (Movement option 2), they get a special exception.")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("For Sleep & Nutrition Habits")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("If habit is NOT checked for a day: Reset to Day 1")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("For Strength Training (Movement option 2)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("If strength checkbox NOT checked for a day BUT 3+ strength sessions completed in that week: Allow unchecked day, stay on current day")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("If strength checkbox NOT checked AND fewer than 3 sessions that week: Reset to Day 1")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Tracks strength sessions in dailyLogs object (60 KB) or separate strengthSessions field")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== MISSING FEATURES IN 25KB ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("9. UNFINISHED FEATURES - MISSING FROM 25 KB")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Critical UI Elements")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Site Bar (Top Navigation): 8 instances in 60 KB, 0 in 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Impact: No visual branding at top, less professional appearance")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: HIGH - Restore from 60 KB")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Challenge Description Card: Missing from 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Impact: User lands on hero page with no explanation of what 40 Thrive is before being asked to join")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB version: Features semi-transparent card with 'Real habits. Real results.' headline and three habit categories with emoji")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: CRITICAL - User reported this is missing")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Modal Auth System: 25 KB has inline forms, 60 KB has polished modal")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Impact: 25 KB auth experience is less polished, forms appear alongside content")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB: Modal slides in from right, cleaner UX with tabs and close button")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: MEDIUM - Improves UX but not blocking")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Scrollable Hero Content: 25 KB static, 60 KB has fixed background with scrollable content")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Impact: User reported 'no unmovable photo' - they expect parallax/fixed background effect")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: MEDIUM - Visual polish")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Progress Tracking")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Progress Tab: Missing from 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("User reported: 'no progress tab'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB features: 40-day visual grid, completion status for each day, stats card (current day, completed days, streak)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Bottom nav in 60 KB: ☀️ Today tab + 📊 Progress tab")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: HIGH - Important for user motivation and tracking")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Strength Training Week Dots: Missing from 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB feature: Shows 7 day-of-week dots indicating which days strength sessions logged")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Needed for: Auto-reset exception logic validation")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: HIGH - Functional requirement for strength training exception")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Authentication")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Magic Link Sign-In: Completely missing from 25 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB has: sendSignInLinkToEmail + signInWithEmailLink functions")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("UX: 'Send me a sign-in link instead' button in login modal")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: LOW - Nice-to-have feature")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Database Efficiency")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("updateDoc vs setDoc: 25 KB uses only setDoc (overwrites entire doc), 60 KB uses updateDoc for partial updates")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Impact: 25 KB less efficient but functional")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Should upgrade to: Import updateDoc, use for field-only updates")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: LOW - Optimization, not blocking")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Polish & Styling")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Gradient backgrounds & backdrop filters: 60 KB more polished")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("User reported: 'title and subtitles are not formatted'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("60 KB has: Better typography, spacing, emoji icons in category boxes")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: MEDIUM - Visual quality")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Success/Status Modals: 60 KB has animated feedback modals")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Examples: 'You're all set!' after signup, 'All habits done today!' after completing tasks")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 1 },
        children: [new TextRun("Priority: LOW - UX enhancement")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      // ========== RECOMMENDATIONS ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("10. RECOMMENDATIONS")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("For Next Build (Merge Strategy)")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Base: Use 60 KB version as foundation (has all features)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Code Quality: Refactor & clean up as done in 25 KB (remove bloat)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Verify: All features from 25 KB plus missing features from 60 KB")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Test: Checkbox toggle, auto-reset, strength training exception, all screens")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Deploy: Push merged version to Netlify as final canonical version")]
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Priority Order")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("CRITICAL: Challenge description card (user visibility issue)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("HIGH: Site bar + Progress tab (core functionality)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("HIGH: Strength training week dots (auto-reset exception logic)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("MEDIUM: Modal auth system + scrollable hero (UX polish)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("LOW: Magic link sign-in + updateDoc optimization + animated modals")]
      }),

      new Paragraph({
        children: [new TextRun("")]
      }),

      new Paragraph({
        children: [new TextRun("END OF INVENTORY")]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/friendly-gifted-wright/mnt/outputs/40Thrive_Project_Inventory.docx", buffer);
  console.log("Document created successfully!");
});

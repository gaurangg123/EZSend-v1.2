# 📧 EZSend-v1.2 – Outlook VBA Macro

This VBA macro automates sending emails from the **Outbox** in Microsoft Outlook with smart delays and cleanup.  

---

## ✨ Features
- ✅ Random delay between each email (1–5 minutes)  
- ✅ Deadline cutoff (default: 11:30 PM) – avoids scheduling past the day  
- ✅ Cleans invalid characters from recipient fields (To, CC, BCC)  
- ✅ Sends emails in **HTML format** (preserves formatting)  
- ✅ Deletes originals from Outbox after scheduling (to prevent duplicates)  
- ✅ Shows helpful message boxes when Outbox is empty or cutoff is reached  

---

## 📂 Macro Code
The main macro is called:  
```vb
SendFreshDraftsWithRandomDelay_CleanedHTML
```

Utility function included:  
```vb
CleanEmail()
```

---

## 🛠️ Setup Instructions

### Step 1: Open Outlook VBA Editor
1. Open **Microsoft Outlook**  
2. Press **`Alt + F11`** to open the VBA editor  
3. In the left pane, expand **Project1 (VbaProject.OTM)**  

---

### Step 2: Insert the Macro
1. Go to **Insert > Module**  
2. Copy–paste the full code into the new module  
3. Save the project (`Ctrl + S`)  

---

### Step 3: Add a Quick Access Button (Optional)
1. In Outlook, right-click the ribbon → **Customize the Ribbon**  
2. Create a new group under **Home** (e.g., “Macros”)  
3. Add the macro `SendFreshDraftsWithRandomDelay_CleanedHTML` to this group  
4. (Optional) Assign an icon for easy access  

---

## ▶️ How to Run
- Place your draft emails in **Outbox**  
- Run the macro:
  - From VBA Editor → Press **F5**  
  - From Outlook Ribbon → Click your assigned button  
- The macro will:
  - Process each draft  
  - Schedule with a randomized delay (1–5 mins each)  
  - Stop once the cutoff (11:30 PM) is reached  
  - Delete the originals after scheduling  

---

## ⚠️ Notes
- To change cutoff time:  
  ```vb
  deadline = Date + TimeValue("23:30:00")
  ```
- To adjust delay range (default: 1–5 minutes):  
  ```vb
  randomDelay = Int((5 - 1 + 1) * Rnd + 1)
  ```
- To **keep originals** instead of deleting, comment/remove:  
  ```vb
  originalMail.Delete
  ```

---

## ✅ Example Success Message
At completion, Outlook shows:  
> **"✅ All fresh drafts sent with random delays before 11:30 PM."**

---

## 📌 Requirements
- Microsoft Outlook (Desktop, Windows)  
- Macros enabled (Trust Center → Macro Settings)  
- Basic familiarity with VBA  

---

# Attachment Saver (Outlook VBA)

This small tool adds two buttons (macros) to Outlook that save all attachments
from selected emails into a folder on your computer.

It is designed for **non‑technical users** and runs only on **Windows + Outlook desktop**.

---

## What the tool does

- Saves all attachments from the emails you select
- Puts them directly into one folder (no sub‑folders per email)
- Avoids duplicate names by adding `_1`, `_2`, … to files

You get two actions:

1. **Save to default folder**  
2. **Save to folder I choose**

---

## Before you start

You need:

- Windows PC
- Outlook desktop (e.g. Microsoft 365 / Outlook 2016 or newer)
- Permission to run **macros** in Outlook

If your company blocks macros, you may need IT to enable them.

---

## Step 1 – Open the VBA editor

1. Open **Outlook**.
2. Press **`ALT` + `F11`** on your keyboard.  
   A new window called **“Microsoft Visual Basic for Applications”** opens.

---

## Step 2 – Import the macro file

1. In the VBA window, click **File → Import File…**
2. Select the file **`AttachmentSaverLocal.bas`** that you received.
3. Click **Open**.  
   A new module named `AttachmentSaverLocal` appears on the left side.

You can now close the VBA window (click the **X** in the top right).

---

## Step 3 – Add a button in Outlook

We will add a button so you do not have to open the macro menu every time.

1. In Outlook, click **File → Options**.
2. Choose **“Quick Access Toolbar”** on the left  
   (or **“Customize Ribbon”** if you prefer a button in the ribbon).
3. In the drop‑down **“Choose commands from”**, select **“Macros”**.
4. You should see:
   - `Project1.AttachmentSaverLocal.SaveAttachments_DefaultFolder`
   - `Project1.AttachmentSaverLocal.SaveAttachments_SaveAs`
5. Select one of them (start with `SaveAttachments_SaveAs`) and click **Add >>**.
6. (Optional) Click **Modify…** to choose an icon and friendly name, e.g.
   - Name: `Save Attachments`
7. Click **OK** to close the Options window.

You now have a button in Outlook that runs the macro.

---

## Step 4 – How to use it

1. In Outlook, select one or several emails in your inbox.
2. Click your new **“Save Attachments”** button (or run the macro via **View → Macros**).
3. If you use the **“SaveAs”** version:
   - A window appears asking for a folder.
   - Choose the folder where attachments should be stored.
4. The tool saves all attachments from the selected emails into this folder.
5. When it is finished, a message box shows:
   - How many attachments were saved
   - How many emails had no attachments
   - If any errors occurred

---

## Default folder (for the “Default” button)

The macro uses this standard folder on your computer:

```text
C:\EmailAttachments
```

When you use **“Save to default folder”**, all files are saved there automatically.

You can change this folder only inside the VBA code.  
If you need that and are not sure how, ask a technical person to adjust the line:

```vb
Public Const DEFAULT_SAVE_PATH As String = "C:\EmailAttachments"
```

---

## Common questions

### I do not see “Macros” in Outlook options

Your Outlook may have macros disabled by your company.  
In that case, ask your IT support to enable **VBA macros in Outlook** for you.

### Is this safe?

The macro only:
- Looks at the emails you select in Outlook
- Saves their attachments as files on your computer

It does not send data anywhere.

---

## How to remove it

If you no longer want to use the tool:

1. Press **`ALT` + `F11`** to open the VBA editor.
2. In the left list, right‑click the module **`AttachmentSaverLocal`**.
3. Click **Remove AttachmentSaverLocal…** and confirm.
4. Remove the macro button from the Quick Access Toolbar / Ribbon:
   - Outlook → **File → Options → Quick Access Toolbar** (or **Customize Ribbon**)
   - Select the macro on the right and click **Remove**.

The tool is then completely removed.

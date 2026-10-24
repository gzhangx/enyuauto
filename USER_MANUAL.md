# Enyu Project Management — User Manual

This site is for the EnYu project manager. It keeps one Excel row per article, links that row to a SharePoint folder, and creates assigned FreedCamp tasks.

**Live site:** https://gzhangx.github.io/enyuauto/

Use Microsoft Edge or Chrome. Sign in with the EnYu Editors Microsoft account that can open the SharePoint library and `enyustatus.xlsx`.

If you are new, read **section 2 (How an article moves)** before you create any task. That is the part people usually have to ask about.

---

## 1. Sign in

1. Open the live site.
2. Click **Sign in with Microsoft**.
3. Use the EnYu Editors account.
4. The **Projects** page loads.

Your name and email appear in the top-right. Task templates use that login: `{签名}` is your Chinese name from the editor sheet, `{sender}` is your English name.

### FreedCamp login (first time, or after credentials change)

1. Open **FreedCamp Login** in the top bar.
2. Enter username, password, and EnYu WP token if you use it.
3. Click **Save to OneDrive**.
4. Return to **Projects** and click **Reload All**.

Credentials are stored in `enyu_secs.xlsx` at the root of your OneDrive.

---

## 2. How an article moves (read this first)

Every article lives in **one SharePoint folder** (`mainFolder`). Inside it, the numbered subfolders must use these exact names:

| Folder | Who works here | What belongs here |
| --- | --- | --- |
| `0. Original Materials` | Incoming manuscript | The author’s original file. Put it here **before** you create 校对. |
| `1 Text Editor & First Proofread` | 校对 (first proof / 一校) | Working copy for the first editor. The site **copies** the file here when you create 校对. |
| `2 Second Proofread` | 二校 | Working copy for the second editor. The site **copies** from folder 1 into folder 2 when you create 二校. |
| `3 Art Editor` | 美编 | Layout / art copy. The site **copies** from folder 2 (or folder 1 if 2 is empty) into folder 3 when you create 美编. |
| `4 Publish` | 发布 | Final signed PDF / Word. Used for the live site, 文章摘要, and 相关文章. |
| `5 Media` | 语审 | Audio (`.mp3`). The site does **not** copy this file; it only puts the link in the task. |

Typical order:

**Put original in folder 0 → create 校对 (copies into folder 1) → editor finishes in folder 1 → create 二校 (copies into folder 2) → create 美编 (copies into folder 3) → put the signed file in folder 4 → create 发布.**

Do not skip the copy. 二校 should not start from the file still sitting only in folder 0. Create 校对 first so folder 1 has the first-proof copy.

### What 文章链接 is (and is not)

**文章链接** is a field on the Excel row. When you **Add New Project**, you may Browse and save a file URL there (often the original article). That is a bookmark on the sheet.

It is **not** what 校对 / 二校 / 美编 / 教审 / 语审 browse when you create a task.

Those actions always look in the numbered folders above:

- Folder **0** and folder **1** are for **校对**, not for 文章链接.
- Folder **1** and folder **2** are for **二校**.
- **文章链接** does not fill folder 0 or folder 1 by itself. You still put the original file into folder 0 in SharePoint.

If the FreedCamp template has `{article}`, the link in the task is usually the **copied** file in the destination folder (folder 1 after 校对, folder 2 after 二校), not the 文章链接 cell.

---

## 3. Add a new project

Do this only after the SharePoint article folder exists and, if you already have the manuscript, after the original file is in `0. Original Materials`.

1. Click **Add New Project**.
2. Fill **文件** in this order, then click outside the box:
   - `{category} {YYYY-MM-DD} {author} {article title}`
   - Example: `宣教无国界 2026-10-21 Jane Lee Who Is that in the Painting`
   - For **佳作分享**, leave the article title off. It is a book passage, so **文章名** can stay empty.
3. Or click **Browse SharePoint**, open `General` → year-month → the article folder, then **Use this folder**.
   - That fills **文件** from the folder name and saves a SharePoint **sharing** link in **mainFolder** (`:f:/s/...`).
4. Check **文章类别**, **slug**, **作者**, **称呼**, and **作者电邮**.
   - **Search** on email uses that author’s last known address from earlier rows.
5. **文章链接** is optional. Use **Browse** if you want to store a source-file URL on the sheet. This does **not** replace putting the file in folder 0.
6. **备注** is optional notes for this 文件. It is stored on the Excel **备注** column.
7. Click **Save**.

**mainFolder** must be the article folder’s sharing link, not a long `/Shared Documents/...` path with the 文件 name stuck on the end.

---

## 4. Create a FreedCamp task

Create **one action at a time**. Click the **empty** cell for that action (校对, 二校, 美编, …). Cells that already have an assignee are not clickable for create.

The task is assigned to the editor you pick. FreedCamp may email them. Remind them on WeChat yourself if they do not check email.

### Fill the dialog

1. **Editor** — required. The table shows open tasks, last assigned, last completed, and total done. Red rows are on leave. Click a row to select that person.
2. **Due in** — days from today (default 7). A warning appears if that due date falls during leave.
3. **English template** — check this for an English article (only if a template exists).
4. **AI template** — use the AI template when it exists. Human-only editors are hidden while this is on.
5. **File** — pick the SharePoint file for this action (see the table below). Click **open** to check it.
6. **FreedCamp preview** — the task description. You can edit it. `{相关文章}` and `{文章摘要}` are filled from `4 Publish` when those placeholders exist.

Then **Create Task** → confirm the task name and assignee → wait for create and **Reload All**. The cell should show the assignee, not Delete.

If the preview is wrong, fix it before you confirm. Do not create a test task in FreedCamp unless you mean to keep it.

### Which file to pick (this is the copy step)

The **File** dropdown lists files from the folders for that action (`/folder/filename`). If a copy already exists in the **destination** folder, the dialog prefers that file. The line says **Will copy** `[file]` **into folder** `/1 Text Editor & First Proofread` (destination folder in bold). **Open File** opens the selected file. Do not look for a separate “from folder” note — the source folder is already in the dropdown.

| You are creating | Dropdown lists | What Create Task does with the file you pick |
| --- | --- | --- |
| **校对** (一校) | `0. Original Materials` **and** `1 Text Editor & First Proofread` | **Copies** it into folder **1**. First time: pick the original in folder 0. Later: you can pick the copy already in folder 1. |
| **二校** | Folder **1** **and** folder **2** | **Copies** it into folder **2**. Normal case: pick the finished first-proof file in folder 1. |
| **美编** | Folder **2**. If that is empty, folder **1** | **Copies** it into folder **3**. |
| **教审** | Folder **0** only | Does **not** copy. The task links to the original in folder 0. |
| **语审** | `.mp3` files in folder **5** | Does **not** copy. The mp3 link goes into `{voice_file}` in the description, not `{article}`. |
| **发布** | Whatever the Excel mapping says (usually `4 Publish`) | Follow the File line in the dialog (`Will copy … into folder /…` vs `File from …`). 文章摘要 / 相关文章 need a signed file in folder 4 (`*_signed.pdf` or the matching Word file). |

**校对 vs 二校, in practice**

1. Manuscript is in folder 0. Folder 1 may still be empty.
2. Create **校对**, pick the folder 0 file. The site copies it to folder 1. The first editor works in folder 1.
3. When first proof is done, the updated file is in folder 1 (the editor saves there).
4. Create **二校**, pick that folder 1 file. The site copies it to folder 2. The second editor works in folder 2.

**文章链接 is not used in these dropdowns.** If folder 0 is empty, 校对 cannot pick a file even if 文章链接 is filled. Put the file in folder 0 first.

If folder 1 is empty when you create **二校**, the File line says **Create 校对 first, or put the first-proof file in folder 1** — not a generic list error.

Other action names on the sheet follow the Excel config `mapping` rows (copy from a folder name, or from a sheet column such as 文章链接). If the File line says `Will copy … into folder /2 Second Proofread.`, creating the task will copy. If it only says `File from …`, it only inserts a link.

---

## 5. Projects page

Each row is one article on the main Excel sheet.

| Control | What it does |
| --- | --- |
| **Search** | Filters by **author**, **article**, **date**, **category**, or **备注**. A search also finds Done rows. |
| **Add New Project** | Appends a new article row |
| **Reload** | Reloads the Excel list only |
| **Reload All** | Reloads Excel and FreedCamp (use after creating a task, or when cells look stale) |
| **Sync All** | Writes FreedCamp values back to Excel when orange sync buttons appear |
| **Show all** | Includes rows marked Done |
| **days** | How many days a completed date stays highlighted, and how soon a due date is treated as urgent (0–10) |

The number in the top-right corner opens a log of recent actions.

### Reading a row

- **Line** — first column, Excel row number. Click the number to **edit this project** (author, folder, slug, 备注, …). Orange **Sync** copies matching FreedCamp values into the sheet. Hover for the field list.
- **文件** — article name. Click it to open the SharePoint folder (`mainFolder`).
  - **preview** appears after 发布 is completed and the row is not Done.
  - **view on website** appears after **Done** is checked.
- **校对 / 二校 / 美工 / 语审 / 教审 / 发布** (and any other actions on the sheet) — that FreedCamp task. Click an **empty** cell to create it.
- **Done** — writes `Y` on the sheet and hides the row unless **Show all** is on.

Click the person icon on an action column header for assignment history.

### What an action cell shows

| You see | Meaning |
| --- | --- |
| Empty cell | No task yet. Hover shows **Click the Cell to Create 校对 Task** (or 二校, …). Click the cell to create it |
| **Delete** plus a task ID | A FreedCamp task exists but is not assigned. You can delete it here |
| Editor name | Task is assigned. The name turns **red** if the task is past due or due within the **days** window (default 3). Hover for the due date. Click the name to open the task. Delete is hidden |
| A date in green or red | Completed in FreedCamp. Recent dates use the **days** highlight. Click the date to open the same task |
| **Done** | Sheet task ID is marked done |
| **No English** | This action is skipped for English-only articles |

Once a task is assigned, do not delete it from this grid. To undo a mistake: delete the task in FreedCamp, clear that action’s `TaskId` cell in Excel (empty, not `done`), then **Reload All**.

### Edit a project / 备注

Add a column named **备注** on the **main** Excel sheet header row — **at the end (rightmost)**, not inserted on the left. Save the workbook, then **Reload**. Inserting on the left shifts 校对 / 美编 / 教审 and the Sync tooltip will say those columns were not found.

Click the **Line** number. The popup title is the **文件** name, with the **备注** under it. Change author, folder, slug, article fields, or the **备注**, then **Save**. A message confirms `{文件} 文件 saved successfully`. The 备注 also shows under **文件** on the grid (two lines). This does not change FreedCamp tasks.

### Edit a FreedCamp task from this site

Click the editor name or the completed date. The popup is an editor, not a copy of the FreedCamp website.

You can change **title**, **status** (No Progress / In Progress / Completed), **assignee**, **due date**, and **description**. Assigned lists editors for that action from the editor sheet, not everyone in FreedCamp. Comments stay read-only.

Click **Save to FreedCamp**. The grid updates that cell without a full reload. Use **Reload All** if another tab or FreedCamp itself changed the task.

---

## 6. Everyday checklist

1. Sign in with Microsoft. **Reload All** if the list is empty or old.
2. Make sure the SharePoint article folder exists, with the six numbered subfolders.
3. Put the original manuscript in `0. Original Materials`.
4. **Add New Project** if the article is not on the sheet yet. Save the folder sharing link as **mainFolder**.
5. Click empty **校对** → pick the folder 0 file → Create Task. The file is copied to folder 1.
6. When first proof is done, click empty **二校** → pick the folder 1 file → Create Task. The file is copied to folder 2.
7. Create **美编** / **语审** / **教审** / **发布** the same way, using the file table in section 4.
8. After each create, the cell shows the assignee. Remind them on WeChat yourself if needed.
9. When the article is finished on the site, check **Done**.

---

The top bar has **Projects** and **FreedCamp Login**. Daily work stays on Projects. Use FreedCamp Login only to save username/password the first time (or when they change).

---

## 7. Troubleshooting

**The page looks like the old version.**  
Hard-refresh (Ctrl+F5). Live site: https://gzhangx.github.io/enyuauto/. Local: http://localhost:5173/enyuauto/.

**Clicking a cell does nothing.**  
That action already has a task ID. Only empty cells open Create Task.

**校对 / 二校 says there is no file.**  
The original is not in folder 0, or first proof is not in folder 1. **文章链接** on the sheet does not count. Open the article folder in SharePoint and check the numbered subfolder names (they must match section 2 exactly).

**I created 二校 but the editor is still looking at folder 1.**  
Create Task copies the file you picked **into folder 2**. Confirm the File dropdown was the folder 1 file, then look in `2 Second Proofread`.

**Delete still shows after I created a task.**  
Wait for the automatic reload, or **Reload All**. If it still shows Delete, FreedCamp did not match the editor and the task was left unassigned.

**I need to undo a task.**  
Delete it in FreedCamp, clear `{action} TaskId` in Excel, **Reload All**.

**The 文件 name opens the wrong place.**  
Paste the SharePoint **Copy link** URL into `mainFolder`, then Reload.

**Related articles or 文章摘要 are missing.**  
Need Microsoft login, `mainFolder`, and a signed article in `4 Publish`. If search returns 400, that is the server search function, not the sheet.

**“Microsoft token not available” or “sign-in expired” / invalid token lifetime.**  
The Microsoft Graph token lasted about an hour. Sign in with Microsoft again from the top right, then click the cell. This is not a missing folder.

**Projects stay empty.**  
Save FreedCamp credentials, confirm you can open `enyustatus.xlsx`, then **Reload All**.

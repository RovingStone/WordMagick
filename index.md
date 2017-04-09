---
layout: default
---
# System requirements
A few:
{:.lheader}
* Office for Windows Desktop 2016 (build 16.0.6769.0000 or later)
* Internet connection

You don't need to sign-in into Word with your Microsoft Account.

# Installation
The following steps are necessary to complete just once. To upgrade the add-in you just need to replace the manifest file with a new one.
## 1. Share a folder
1. On the Windows computer go to the drive letter, or the folder you want to use as your shared folder catalog.
2. Open the context menu for the folder (right-click) and choose **Properties**.
3. Open the **Sharing** tab.
4. On the **Choose people ...** page, add yourself and anyone else with whom you want to share add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.
5. Choose **Share**.
6. In the next window open the context menu for the shared element (right-click) and choose **Copy link**.
7. Choose **Done** > **Close**.

## 2. Prepare the manifest
1. Choose "ZIP" or "7z" above to download the manifest archive.
2. Unpack the archive, move WordMagick.xml manifest into folder, shared on the previous step.

## 3. Specify the shared folder as a trusted catalog
1. Open a new document in Word.
2. Choose the **File** tab, and then choose **Options**.
3. Choose **Trust Center**, and then choose the  **Trust Center Settings** button.
4. Choose  **Trusted Add-in Catalogs**.
5. In the  **Catalog Url** box, enter the full network path to the shared folder catalog. Press <kbd>Ctrl-V</kbd> to insert copied shared folder link. Network path is a line after "file:", so delete the rest. You will get something like this: "//COMPNAME/Folder".
6. Choose **Add Catalog**.
7. Select the **Show in Menu** check box, and then choose **OK**.
8. Close the Word application so your changes could take effect.

## 4. Load Word Magick
1. In Word select **My Add-ins** on the **Insert** tab of the ribbon.
2. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.
3. Select the **Word Magick** add-in and choose OK to insert the add-in.
4. Word Magick would be placed on the **General** tab of ribbon.

# Usage
Select some words or a text, or some table cells, click on any of Word Magick buttons, and let the magic begins. You can select the text, or some table cells, but not both (if you don't won't to create a text instead of a table). You can work this way only with the English alphabeth. All other letters will be treated as non-word letters (they are considerd as the punctuation by the Word Magick program).

<div class="stats" markdown="1">
[![GitHub release](https://img.shields.io/github/release/RovingStone/WordMagick.svg)](https://github.com/RovingStone/WordMagick/releases/latest) [![GitHub release](https://img.shields.io/github/tag/RovingStone/WordMagick.svg)](https://github.com/RovingStone/WordMagick/releases/tag/v0.1.0) [![license](https://img.shields.io/github/license/RovingStone/WordMagick.svg)](https://github.com/RovingStone/WordMagick/blob/master/LICENSE) 
</div>

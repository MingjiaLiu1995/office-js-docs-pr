---
title: Implement event-based activation in WXP add-ins (preview)
description: Learn how to develop a WXP add-in that implements event-based activation.
ms.date: 08/01/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement event-based activation in WXP add-ins (preview)

With the feature, develop an add-in to automatically activate and complete operations when certain events occur in Word, Excel, and PowerPoint, such as create new documents and open existing documents. 

> [!NOTE]
> The feature is in preview for early testing. You can not deploy a add-in with this feature to your customer yet. Please also notice the preview version of feature may be different from the released version.


## Supported events and clients

| Event name | Description | Supported clients and channels |
| ----- | ----- | ----- |
| `OnDocumentOpen` | Occurs on a WXP document opens (includes create a new document and open an existing document).| <ul><li> [To be updated] Word desktop? Word WAC?? Excel? PowerPoint?? </li></ul> |

The following sections walk you through how to develop a Word add-in that automatically changes the document header when a new or existing document opens. This highlights a sample scenario of how you can implement event-based activation in WXP add-ins.

## Set up your environment

To run the feature, you must have a supported version of Word and a Microsoft 365 subscription. Then, create a Word add-in project.

## Configure the manifest

[Instruction and code samples]

## Implement the event handler

[Instruction and code samples]

## Add a reference to the event-handling JavaScript file

[Instruction and code samples]

## Test and validate your add-in

[Instruction and code samples]

## Behavior and limitations

[Instruction and code samples]

Currently, the feature is only supported in add-in only manifest, please contact product team if you need unified app manifest.

[More to be added]

## Additional supported APIs

[Instruction and code samples]

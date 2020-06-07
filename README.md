# MS_Word_HiddenStorage
MS_Word_HiddenStorage provides the ability save metadata in docx files.

## Introduction
Sometimes, when developing addins for MS Word, you may want to store some metadata outside of main document, invisible to users.
In MS Excel it is a trivial task with hidden sheets, but MS Work does not have anything like it. So this library aims to provide a solution.
It only works for .docx files.

## How it works
MS Word .docx files have CustomXMLParts that allow you to attach any object as XML. This library helps to access by key and manage these attached parts.
The library can only work with strings, so every object must be serialized into a string to store it.

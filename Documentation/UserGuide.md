# Biologix User Guide

Welcome to Biologix, a Visual Basic program designed for managing invoices in a fictional chemical company. This User Guide provides an overview of the program's features and functionalities.

## Table of Contents
1. [Introduction](#introduction)
2. [Login Screen](#login-screen)
3. [Password Reset](#password-reset)
4. [Main Form](#main-form)
5. [Loan Calculator](#loan-calculator)
6. [Web Browser](#web-browser)
7. [Memo Notepad](#memo-notepad)
8. [Extra Features](#extra-features)

## 1. Introduction <a name="introduction"></a>

Biologix is a Visual Basic program developed in 2015 during college. It serves as an invoicing solution for a fictional chemical company. The program features a login screen, a main form for creating invoices, a loan calculator, a web browser, a memo notepad, and additional features to showcase coding skills in Visual Basic.

## 2. Login Screen <a name="login-screen"></a>

The login screen requires a username and password. The default credentials are "admin" for both. Password reset functionality is available on the main form. Note that security considerations for username/password storage are discussed in the [Security Update](#password-reset).

![Login Screen](URL_TO_IMAGE)

## 3. Password Reset <a name="password-reset"></a>

On the main form, users can reset their password. The password is saved as a sys.dll file for encryption. However, it's important to note that this method is considered insecure. See the [Security Update](#password-reset) for more details.

![Password Reset](URL_TO_IMAGE)

## 4. Main Form <a name="main-form"></a>

The main form allows users to input price, quantity, product, and product ID for invoices. The program calculates the price of the product based on quantity. Users can save forms, product prices, and totals. The program can save over 100 products and includes error handling.

![Main Form 1](URL_TO_IMAGE)
![Main Form 2](URL_TO_IMAGE)
![Main Form 3](URL_TO_IMAGE)

## 5. Loan Calculator <a name="loan-calculator"></a>

The loan calculator helps users calculate payments for larger invoice items, facilitating easier payment planning.

![Loan Calculator](URL_TO_IMAGE)

## 6. Web Browser <a name="web-browser"></a>

The web browser allows users to research products or access the main website for more information. It includes a progress bar and a built-in contact button for emailing the company.

![Web Browser 1](URL_TO_IMAGE)
![Web Browser 2](URL_TO_IMAGE)

## 7. Memo Notepad <a name="memo-notepad"></a>

The Memo Notepad is a program for taking and saving notes in Rich Text or Text files.

![Memo Notepad 1](URL_TO_IMAGE)
![Memo Notepad 2](URL_TO_IMAGE)

## 8. Extra Features <a name="extra-features"></a>

Additional features include an About page, Disclaimer, edit tabs and shortcuts, and a help feature for the program.

![Extra Features](URL_TO_IMAGE)

## Security Update

It's important to note that the program's security has some flaws. Passwords are stored in a file and considered insecure. Future programs should adopt more secure practices.

## Folder Structure on GitHub

To maintain a clean structure on GitHub, the source code files can be organized as follows:
```
Biologix/
│
├── Documentation/
│   ├── UserGuide.md
│   └── Screenshots/
│       ├── LoginScreen.png
│       ├── MainProgram.png
│       ├── MemoPad.png
│       └── ... (other screenshots)
│
├── SourceCode/
│   ├── Calculator/
│   │   └── Calculator.vb
│   ├── Loops/
│   │   └── Loops.vb
│   ├── CardGame/
│   │   └── CardGame.vb
│   ├── Array/
│   │   └── Array.vb
│   ├── Fractals/
│   │   └── Fractals.vb
│   ├── Functions/
│   │   └── Functions.vb
│   ├── LoginScreen/
│   │   └── LoginScreen.vb
│   ├── MainProgram/
│   │   └── MainProgram.vb
│   ├── MemoPad/
│   │   └── MemoPad.vb
│   ├── Module/
│   │   └── Module.vb
│   ├── Pong/
│   │   └── Pong.vb
│   └── Animation/
│       └── Animation.vb
│
└── README.md
```

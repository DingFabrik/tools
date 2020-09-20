# Spendenquittungen erstellen mit LaTeX und Python

Eines der Mitglieder der Dingfabrik, Uwe Ziegenhagen (https://www.uweziegenhagen.de) hat auf der Basis von LaTeX-Vorlagen einen Workflow zum Erzeugen von Spendenquittungen entwickelt.

Was wird benötigt, um die Dokumente zu erzeugen:

* Eine LaTeX-Installation, vorzugsweise TeX Live in einer halbwegs aktuellen Version
* Python mit einigen Paketen (pandas, etc.)
* Excel-Dateien mit Stammdaten und Transaktionsdaten, siehe die Beispieldateien. Als Transaktionsdaten nutzt die Dingfabrik Excel-Exporte aus Quicken, die Logik lässt sich aber leicht adaptieren.

## LaTeX-Vorlage

Für normale Geldspenden befindet sich die Vorlage in der Datei 'Sammelbestaetigung_Geldzuwendung.tex', das Ergebnis der LaTeX-Übersetzung in der Datei 'Sammelbestaetigung_Geldzuwendung.pdf'. Weitere Vorlagen finden sich im folgenden Repo: https://github.com/UweZiegenhagen/spendenquittungen-mit-latex


## Python-Installation

Das ganze Projekt wurde auf der Basis einer Anaconda Python-Installation entwickelt, daher empfehlen wir auch diese.

## Python-Code

Die Datei 'ErzeugeSpendenquittungen.py' enthält den vollständigen Code für die Erzeugung der PDF-Dateien. 

Es werden erst die Stammdaten aus Excel gelesen, anschließend erfolgt das Lesen der Transaktionen. Für jeden gültigen Stammdatensatz (Mitgliedsart 'O' oder 'A') wird eine Quittung erzeugt. 

Die LaTeX-Vorlage wird eingelesen und mittels Jinja2 befüllt. Jinja2 ist eine Template-Engine, die eine schöne Kapselung von LaTeX- und Python-Code erlaubt.

Hinweis: Die TeX-Dateien und PDFs werden in einem 'fertig' Unterordner gespeichert, dieser muss als existieren!


English summary: This repository contains the LaTeX templates and Python scripts to create donation receipts for German nonprofit organizations.



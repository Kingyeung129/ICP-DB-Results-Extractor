# Documentation

## Introduction

This repo contains a python script to extract ICP results data and remove irrelevant testing data. Irrelevant data is defined to be tests done for XINSHA. The results database will be duplicated for backup purposes and XINSHA ICP tests data will be removed after running this script.

## Installation

### Prerequisites

This script is written and tested in Python version 3.12. Please install this version of python before proceeding. Microsoft Access Driver has to be installed as well.

### Install python packages

<pre><code>pip install -r requirements.txt</code></pre>

## Running the script

<pre><code>python script.py -f {Target file path of Microsoft Access Database to extract data from}</code></pre>

## Graphical User Interface (GUI)

A graphical user interface has been developed using python tkinter library. Code has been compiled using auto-py-to-exe into an windows executable. Do note that GUI is only supported for Windows OS platform.

## Developments

This script will be compiled as an executable or rewritten in batch or powershell for compatibility.
# spe-simple: A Simple SharePoint Embedded Example

This is a simple example to show off SharePoint Embedded in action.

## Pre-reqs
* read https://aka.ms/start-spe
* git
* node.js
* M365 tenant
* VS Code

## Step 1
* Install VS Code SharePoint Embedded extension
* create your trial container type (max 5, life span 30 days) - use a standard container type for any real work
* export your postman configuration document

## Step 2
* Clone this repo
* cd /spe-simple
* npm install
* copy .env_template to .env
* copy your values from postman configuration to .env
* node client.js

## Step 3 (in client.js)
* Get an access token
* Create a container
* Set your active container
* Upload the sample file
* Edit it in Word



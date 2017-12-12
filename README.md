# Telegram chatbot for Qlik Sense

This is example source code for QlikBotNet, a Qlik Sense chatbot for use with the Telegram messaging platform.
QlikBotNet is the middle layer between Telegram and Qlik products which enable users to perform analytics from any devices.

![Bot Basics](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Basics.PNG)

## Current Feature List
* Natural Language Query
* Get KPI values
* List Measures and get current values
* List Dimensions and get KPIs based on selected dimension
* Get related charts
* Get Reports
* Speak in English, Spanish, Portuguese, Italian, Russian and French

## Planned Feature List
* Set Alarms
* Get related news
* Integrate Natural Language Generation
* Port to other messaging platforms

# Getting Started

## For Users

### Basic Usage
| Command | Response | Notes
| ------- | -------- | ---- |
| Help | Shows available commands | |
| Kpi | Shows some buttons with Measures from the master items. At the end, the Bot will show an Analysis button with access to the app in the Qlik Sense Hub. | This is a way to show a quick access to the most used measures. And to jump to the hub to have all the analytic capabilities from Qlik Sense. The last used measures will appear first. |
| Kpi &lt;measure&gt; | Shows the value of measure | Current filters are applied. | 
| Reports | Shows all available NPrinting reports to the user. Then the user will be able to download any of them, and open it directly in Telegram. | The Bot will show all the PDF files available in the Report folder. |
| Dimensions | Shows with buttons all the dimensions included in master items. When one of them is selected, the Bot sends a message with an analysis of the dimension by the last measure used. | The Bot maintains the context during the conversation, and uses the last measure asked. |
| Measures | Shows with buttons all the measures included in master items. When one of them is selected, the Bot sends a message with its value. | It shows the total value for this measure, with no filters. Change App The Bot shows the applications published in the Qlik Sense Streams defined in the config file. When the user selects one app, the Bot will close the current one and open the new one. Every user could be connected to different apps. |
| Language | The Bot will show buttons for the user to select the desired language to receive the Bot messages. | If any user changes the language, at this moment it will be changed for all users. |
| English | The language is changed to English |
| Español | The language is changed to Spanish |
| Português | The language is changed to Portuguese |
| Italiano | The language is changed to Italian |
| Pусский | The language is changed to Russian |
| Français | The language is changed to French |
| /demo | Switch the demo mode (on or off) | When the Demo mode is activated, the Bot will only listen to the Bot Administrator |
| Clear &lt;dimension&gt; | Remove the filter applied by default to that dimension. |
| Clear | Remove all filters |

![Detailed User Guide](https://github.com/qlik-bots/QlikBotNet/blob/master/Documentation/User%20guide%20for%20Telegram%20Bot.V1.0.pdf)

## For Administrators
![Bot Architecture](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Architecture.PNG)

![Detailed Installation Guide](https://github.com/qlik-bots/QlikBotNet/blob/master/Documentation/Installation%20guide%20for%20Telegram%20Bot.V2.1%20-%20Header.pdf)

## For Developers
For developers, please make sure you have completed all of the steps at ![Open Source at Qlik](https://github.com/qlik-bots/open-source)

![Bot Modules](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Modules.PNG)

![Developer Wiki](https://github.com/qlik-bots/QlikBotNet/wiki)

## Prerequisites
This Bot is built with .Net framework, make sure all of the Nuget packages are correctly referenced. Below is a class diagram describing the dependencies between all classes.


## Tips on modifying QlikSense chatbot
* Make telegram chatbot smarter
	* You will need to make several modification
		* BotOnMessageReceived (QlikSenseBot.cs line 513)
			* Taking care of messages received from telegram
		* ProcessConversationResponse (QlikSenseBot.cs line 852)
			* Decide what to do with response from ConversationService
		* BotOnCallbackQueryReceived (QlikSenseBot.cs line 979)
			* Taking care of call back query from telegram
		* Conversation.cs 
			* General Qlik sense related operations
		* New NLP service
			* This is relatively easier, you just need to create your NLP helper class and modify QlikNLP.cs to use it.
* Making a chatbot for slack
	* You will need to create a QlikSlackBot to replace QlikSenseBot 
	* You will need to rewrite user management functions in ConversationService

## Refactoring (In progress)
	
* The chatbot is also undergoing a process of reconstruction, the objective of refactoring is to
	* Enhance code reusability
	* Enhance readability
	* Enhance scalability
	* Enhance maintainability
	* More reliable
	* Estimated day of completion
		* It is an ongoing project so there is no fixed deadline
	* Your help is welcome, please join our development team on [Qlik Community](https://community.qlik.com/groups/qlik-chatbots) or on [Qlik Branch Slack](http://branch.qlik.com/slack)

## Acknowledgments
*Credit and special thanks to **Juan Gerardo Cabeza** and others for all of their hard work in version 1.0 and 2.1.*

## Copyright
Copyright 2017 QlikTech International AB

Licensed under the Apache License, Version 2.0 (the "License");you may not use this file except in compliance with the License.You may obtain a copy of the License at    

![http://www.apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, softwaredistributed under the License is distributed on an "AS IS" BASIS,WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.See the License for the specific language governing permissions andlimitations under the License.


# Telegram chat bot for Qlik sense

This is the source code for QlikBotNet, a Qlik sense telegram chatbot.
QlikBotNet is the middle layer between Telegram and Qlik sense products.
It provides advance conversational analytics to Qlik sense user.
![alt text](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Basics.PNG)
It serves as the middle layer of all advance platforms. Enable users to perform analytics easily on any devices.
![alt text](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Architecture.PNG)
![alt text](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/Bot%20Modules.PNG)

## Note to developers
This Bot is built with .Net framework, make sure all of the Nuget packages are correctly referenced. Below is a class diagram describing the dependencies between all classes.
![alt text](https://raw.githubusercontent.com/qlik-bots/QlikBotNet/master/Documentation/class%diagram.PNG)

## Tips on modifying QlikSense chatbot
Markup:	*Making a chatbot for slack
			*You will need to create a QlikSlackBot to replace QlikSenseBot 
			*You will need to rewrite user management functions in ConversationService
		*Make telegram chatbot smarter
			*You will need to make several modification
				*BotOnMessageReceived (QlikSenseBot.cs line 513)
					*Taking care of messages received from telegram
				*ProcessConversationResponse (QlikSenseBot.cs line 852)
					*Decide what to do with response from ConversationService
				*BotOnCallbackQueryReceived (QlikSenseBot.cs line 979)
					*Taking care of call back query from telegram
				*Conversation.cs 
					*General Qlik sense related operations
		*New NLP service
			*This is relatively easier, you just need to create your NLP helper class and modify QlikNLP.cs to use it.

## Refactoring (In progress)
	
Markup:	*The chatbot is also undergoing a process of reconstruction, the objective of refactoring is to
			*Enhance code reusability
			*Enhance readability
			*Enhance scalability
			*Enhance maintainability
			*More reliable
			*Estimated day of completion
				*It is an ongoing project so there is no fix deadline, but the first release will be available by the end of 2017
			*Any help is welcome, please leave a comment

## Acknowledgments
*Credit and special thanks to **Juan Gerardo Cabeza** and others for all of their hard work in version 1.0 and 2.1.

Copyright 2017 QlikTech International AB
Licensed under the Apache License, Version 2.0 (the "License");you may not use this file except in compliance with the License.You may obtain a copy of the License at    
http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, softwaredistributed under the License is distributed on an "AS IS" BASIS,WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.See the License for the specific language governing permissions andlimitations under the License.


var SERVER_URL = "http://ec2-13-51-0-115.eu-north-1.compute.amazonaws.com:8080";

function onHomepage(e) {
  console.log("onHomepage called");
  return createHomepageCard();
}

function createHomepageCard() {
  console.log("Creating homepage card");
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("OttoMate"));
  
  // Create tabs
  var addToQueueTab = CardService.newCardSection();
  addToQueueTab.addWidget(CardService.newTextButton()
    .setText("Add to Queue")
    .setOnClickAction(CardService.newAction().setFunctionName("addToOttoQueue")));
  
  var savedEmailsTab = createSavedEmailsSection();
  
  var fixedFooter = CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText("Refresh")
      .setOnClickAction(CardService.newAction().setFunctionName("onHomepage")));

  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setText("Welcome to OttoMate")
      .setWrapText(true)));

  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("Add to Queue")
        .setOnClickAction(CardService.newAction().setFunctionName("showAddToQueueTab")))
      .addButton(CardService.newTextButton()
        .setText("Saved Emails")
        .setOnClickAction(CardService.newAction().setFunctionName("showSavedEmailsTab")))));

  card.addSection(addToQueueTab);
  card.addSection(savedEmailsTab);
  
  card.setFixedFooter(fixedFooter);

  return card.build();
}

function createSavedEmailsSection() {
  console.log("Creating saved emails section");
  var savedEmailsSection = CardService.newCardSection()
    .setHeader("Saved Emails");
  
  try {
    var savedEmails = fetchSavedEmails();
    console.log("Fetched saved emails:", savedEmails);
    if (savedEmails && savedEmails.length > 0) {
      savedEmails.forEach(function(email) {
        savedEmailsSection.addWidget(CardService.newKeyValue()
          .setTopLabel(email.subject || 'No Subject')
          .setContent(email.from || 'Unknown Sender')
          .setBottomLabel(email.timestamp ? new Date(email.timestamp).toLocaleString() : 'Unknown Date')
          .setButton(CardService.newTextButton()
            .setText("View")
            .setOnClickAction(CardService.newAction()
              .setFunctionName("viewSavedEmail")
              .setParameters({emailId: email.id.toString()}))));
      });
    } else {
      savedEmailsSection.addWidget(CardService.newTextParagraph().setText("No saved emails found."));
    }
  } catch (error) {
    console.error('Error in createSavedEmailsSection:', error);
    savedEmailsSection.addWidget(CardService.newTextParagraph().setText("Error loading saved emails. Please try again later."));
  }
  
  return savedEmailsSection;
}

function previewEmail(e) {
  var emailId = e.parameters.emailId;
  var email = GmailApp.getMessageById(emailId);
  
  if (!email) {
    return createErrorCard("Unable to retrieve email details.");
  }
  
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Email Preview"));
  
  var previewSection = CardService.newCardSection();
  previewSection.addWidget(CardService.newKeyValue()
    .setTopLabel("From")
    .setContent(email.getFrom()));
  previewSection.addWidget(CardService.newKeyValue()
    .setTopLabel("Subject")
    .setContent(email.getSubject()));
  previewSection.addWidget(CardService.newKeyValue()
    .setTopLabel("Date")
    .setContent(email.getDate().toLocaleString()));
  
  var emailBody = email.getPlainBody();
  previewSection.addWidget(CardService.newTextParagraph()
    .setText(emailBody.substring(0, 500) + "..."));
  
  card.addSection(previewSection);
  
  // Add AI response section
  var aiResponse = callOpenAI(emailBody);
  var aiSection = CardService.newCardSection()
    .setHeader("AI Assistant");
  aiSection.addWidget(CardService.newTextParagraph()
    .setText(aiResponse));
  
  card.addSection(aiSection);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

function fetchSavedEmails() {
  var userEmail = Session.getActiveUser().getEmail();
  var url = SERVER_URL + "/email";
  
  console.log("Fetching saved emails for user:", userEmail);
  
  try {
    var options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'headers': {
        'user-email': userEmail
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();
    
    console.log('Fetch response - Status Code:', statusCode, 'Content:', contentText);
    
    if (statusCode === 200) {
      var emails = JSON.parse(contentText);
      console.log('Parsed emails:', emails);
      return emails;
    } else {
      console.error('Error fetching emails. Status code:', statusCode, 'Response:', contentText);
      return [];
    }
  } catch (error) {
    console.error('Error in fetchSavedEmails:', error);
    return [];
  }
}

function addToOttoQueue(e) {
  console.log("addToOttoQueue called", e);
  var threadId = e.gmail ? e.gmail.threadId : 
                 e.messageMetadata ? e.messageMetadata.threadId : 
                 e.parameters ? e.parameters.threadId : null;
  
  if (!threadId) {
    console.error("No thread ID found");
    return createErrorCard("Unable to retrieve thread ID");
  }

  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      throw new Error('Unable to retrieve thread');
    }

    var firstMessage = thread.getMessages()[0]; // Get the first message in the thread
    var messageId = firstMessage.getId();
    var emailBody = firstMessage.getPlainBody();

    var emailDetails = {
      subject: firstMessage.getSubject(),
      from: firstMessage.getFrom(),
      timestamp: firstMessage.getDate().toISOString(),
      body: emailBody.substring(0, 1000),
      user: Session.getActiveUser().getEmail(),
      status: "new",
      messageId: messageId
    };

    // Process with OpenAI assistant
    var aiResponseCard = callOpenAI(emailBody);
    var jsonResponse = extractJsonFromAiResponse(aiResponseCard);

    // Add the AI response to the emailDetails
    emailDetails.assistantResponse = JSON.stringify(jsonResponse);

    console.log("Email details with assistant response:", JSON.stringify(emailDetails, null, 2));

    // Send email details and AI response to the server
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(emailDetails),
      'muteHttpExceptions': true,
      'validateHttpsCertificates': false,
      'timeout': 10000 // 10 seconds timeout
    };

    console.log("Sending request to server:", SERVER_URL + "/email", "with options:", JSON.stringify(options, null, 2));

    var response = UrlFetchApp.fetch(SERVER_URL + "/email", options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    console.log("Server response code:", responseCode);
    console.log("Server response body:", responseBody);

    if (responseCode !== 200) {
      throw new Error('Unexpected response from server: ' + responseCode + ' - ' + responseBody);
    }

    // Return the AI response card to display to the user
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(aiResponseCard))
      .build();

  } catch (error) {
    console.error("Error in addToOttoQueue:", error);
    var errorMessage = "Failed to process email. ";
    if (error.message.includes("Address unavailable")) {
      errorMessage += "The server is currently unreachable. Please try again later or contact support.";
    } else {
      errorMessage += error.message;
    }
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(createErrorCard(errorMessage)))
      .build();
  }
}

function extractJsonFromAiResponse(aiResponseCard) {
  var jsonResponse = {};
  
  if (aiResponseCard && aiResponseCard.sections) {
    aiResponseCard.sections.forEach(function(section) {
      if (section.widgets) {
        section.widgets.forEach(function(widget) {
          if (widget.textInput) {
            var key = widget.textInput.name;
            var value = widget.textInput.value;
            jsonResponse[key] = value;
          }
        });
      }
    });
  } else {
    console.error('Invalid aiResponseCard structure:', JSON.stringify(aiResponseCard, null, 2));
  }
  
  console.log('Extracted JSON response:', JSON.stringify(jsonResponse, null, 2));
  return jsonResponse;
}

function onGmailMessage(e) {
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("OttoMate Actions"));

  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextButton()
    .setText("Add to Queue")
    .setOnClickAction(CardService.newAction().setFunctionName("addToOttoQueue").setParameters({threadId: e.gmail.threadId})));

  card.addSection(section);

  return card.build();
}

function onGmailMessageOpen(e) {
  return onGmailMessage(e);
}

function createSuccessCard(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .setNavigation(CardService.newNavigation().updateCard(createHomepageCard()))
    .build();
}

function createErrorCard(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

function callOpenAI(emailBody) {
  var openaiApiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  var assistantId = PropertiesService.getScriptProperties().getProperty('OPENAI_ASSISTANT_ID');
  
  try {
    // Create a thread
    var threadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads', {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + openaiApiKey,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      'muteHttpExceptions': true
    });
    
    if (threadResponse.getResponseCode() !== 200) {
      throw new Error('Failed to create thread. Status: ' + threadResponse.getResponseCode() + ', Body: ' + threadResponse.getContentText());
    }
    
    var threadData = JSON.parse(threadResponse.getContentText());
    var threadId = threadData.id;
    
    // Add a message to the thread
    var messageResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads/' + threadId + '/messages', {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + openaiApiKey,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      'payload': JSON.stringify({
        'role': 'user',
        'content': 'Please analyze this email, find the relevant information and respond with a JSON: ' + emailBody
      }),
      'muteHttpExceptions': true
    });
    
    if (messageResponse.getResponseCode() !== 200) {
      throw new Error('Failed to add message. Status: ' + messageResponse.getResponseCode() + ', Body: ' + messageResponse.getContentText());
    }
    
    // Run the assistant
    var runResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads/' + threadId + '/runs', {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + openaiApiKey,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      'payload': JSON.stringify({
        'assistant_id': assistantId,
        'model': 'gpt-4o-mini'
      }),
      'muteHttpExceptions': true
    });
    
    if (runResponse.getResponseCode() !== 200) {
      throw new Error('Failed to run assistant. Status: ' + runResponse.getResponseCode() + ', Body: ' + runResponse.getContentText());
    }
    
    var runData = JSON.parse(runResponse.getContentText());
    var runId = runData.id;
    
    // Poll for completion
    var status = 'in_progress';
    var maxAttempts = 30; // Maximum number of attempts (30 seconds)
    var attempts = 0;
    while ((status === 'in_progress' || status === 'queued') && attempts < maxAttempts) {
      Utilities.sleep(1000); // Wait for 1 second before checking again
      var checkResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads/' + threadId + '/runs/' + runId, {
        'method': 'get',
        'headers': {
          'Authorization': 'Bearer ' + openaiApiKey,
          'OpenAI-Beta': 'assistants=v2'
        },
        'muteHttpExceptions': true
      });
      
      if (checkResponse.getResponseCode() !== 200) {
        throw new Error('Failed to check run status. Status: ' + checkResponse.getResponseCode() + ', Body: ' + checkResponse.getContentText());
      }
      
      var checkData = JSON.parse(checkResponse.getContentText());
      status = checkData.status;
      attempts++;
    }
    
    if (status === 'completed') {
      // Retrieve the assistant's response
      var messagesResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads/' + threadId + '/messages', {
        'method': 'get',
        'headers': {
          'Authorization': 'Bearer ' + openaiApiKey,
          'OpenAI-Beta': 'assistants=v2'
        },
        'muteHttpExceptions': true
      });
      
      if (messagesResponse.getResponseCode() !== 200) {
        throw new Error('Failed to retrieve messages. Status: ' + messagesResponse.getResponseCode() + ', Body: ' + messagesResponse.getContentText());
      }
      
      var messagesData = JSON.parse(messagesResponse.getContentText());
      var assistantMessage = messagesData.data.find(message => message.role === 'assistant');
      
      if (assistantMessage && assistantMessage.content && assistantMessage.content[0] && assistantMessage.content[0].text) {
        try {
          var jsonResponse = JSON.parse(assistantMessage.content[0].text.value);
          return createEditableResponseCard(jsonResponse);
        } catch (parseError) {
          console.error('Error parsing OpenAI response:', parseError);
          console.log('Raw response:', assistantMessage.content[0].text.value);
          return createEditableResponseCard({ error: "Unable to parse AI response. Raw response: " + assistantMessage.content[0].text.value });
        }
      } else {
        console.error('Unexpected assistant message structure:', assistantMessage);
        return createErrorCard('Error: Unexpected response structure from assistant.');
      }
    } else {
      return createErrorCard('Assistant run failed or timed out. Final status: ' + status);
    }
  } catch (error) {
    console.error('Error in callOpenAI:', error);
    return createErrorCard('Error: Unable to process email. Details: ' + error.message);
  }
}

function createEditableResponseCard(jsonResponse) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("OttoMate Assistant Response"));

  if (typeof jsonResponse !== 'object' || jsonResponse === null) {
    console.error('Invalid jsonResponse:', jsonResponse);
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText("Error: Received an invalid response from the AI assistant.")));
    return card.build();
  }

  Object.keys(jsonResponse).forEach(function(key) {
    var section = CardService.newCardSection().setHeader(formatHeader(key));
    
    var value = jsonResponse[key];
    var stringValue = (value !== null && value !== undefined) ? value.toString() : '';
    
    var textInput = CardService.newTextInput()
      .setFieldName(key)
      .setValue(stringValue)
      .setMultiline(key === 'notes' || stringValue.length > 50)  // Make 'notes' or long text multiline
      .setTitle('Edit ' + formatHeader(key));
    
    section.addWidget(textInput);
    card.addSection(section);
  });

  // Add a button to save changes
  var buttonSection = CardService.newCardSection();
  buttonSection.addWidget(CardService.newTextButton()
    .setText("Save Changes")
    .setOnClickAction(CardService.newAction().setFunctionName("saveChanges")));
  card.addSection(buttonSection);

  return card.build();
}

function formatHeader(key) {
  return key.split(/(?=[A-Z])/).join(' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
}

function saveChanges(e) {
  var updatedData = {};
  Object.keys(e.formInput).forEach(function(key) {
    updatedData[key] = e.formInput[key];
  });

  // Here you can add logic to save the updated data to your backend or perform any other necessary actions
  console.log('Updated data:', updatedData);

  // Create a confirmation card
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Changes Saved"));
  
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText("Your changes have been saved successfully."));
  card.addSection(section);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

function checkServerReachability() {
  var url = SERVER_URL + "/health"; // Assuming you have a health check endpoint
  
  try {
    var response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'muteHttpExceptions': true,
      'validateHttpsCertificates': false,
      'timeout': 10000 // 10 seconds timeout
    });
    
    var responseCode = response.getResponseCode();
    var contentText = response.getContentText();
    
    console.log('Server check - Status Code:', responseCode, 'Content:', contentText);
    
    if (responseCode === 200) {
      return "Server is reachable. Status code: " + responseCode;
    } else {
      return "Server responded with status code: " + responseCode + ". Content: " + contentText;
    }
  } catch (error) {
    console.error('Error checking server:', error);
    return "Error: " + error.toString() + ". Message: " + (error.message || "No additional message");
  }
}

function onOpen(e) {
  var ui = CardService.newCardBuilder();
  ui.setHeader(CardService.newCardHeader().setTitle("Server Check"));
  
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextButton()
    .setText("Check Server")
    .setOnClickAction(CardService.newAction().setFunctionName("displayServerStatus")));
  
  ui.addSection(section);
  
  return ui.build();
}

function displayServerStatus(e) {
  var status = checkServerReachability();
  
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Server Status"));
  
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(status));
  
  card.addSection(section);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

function createErrorCard(message) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Error"));
  
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(message));
  
  card.addSection(section);
  
  return card.build();
}

function showAddToQueueTab(e) {
  var card = createHomepageCard();
  card.sections[2].setVisible(true);
  card.sections[3].setVisible(false);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function showSavedEmailsTab(e) {
  var card = createHomepageCard();
  card.sections[2].setVisible(false);
  card.sections[3].setVisible(true);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function viewSavedEmail(e) {
  var emailId = e.parameters.emailId;
  
  try {
    var email = fetchSavedEmailById(emailId);
    if (!email) {
      return createErrorCard("Unable to retrieve email details.");
    }
    
    var card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader().setTitle("Saved Email Details"));
    
    var detailsSection = CardService.newCardSection();
    detailsSection.addWidget(CardService.newKeyValue()
      .setTopLabel("From")
      .setContent(email.from));
    detailsSection.addWidget(CardService.newKeyValue()
      .setTopLabel("Subject")
      .setContent(email.subject));
    detailsSection.addWidget(CardService.newKeyValue()
      .setTopLabel("Date")
      .setContent(new Date(email.timestamp).toLocaleString()));
    detailsSection.addWidget(CardService.newTextParagraph()
      .setText(email.body));
    
    card.addSection(detailsSection);
    
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build()))
      .build();
  } catch (error) {
    console.error('Error in viewSavedEmail:', error);
    return createErrorCard("Error viewing saved email. Please try again later.");
  }
}

function fetchSavedEmailById(emailId) {
  var userEmail = Session.getActiveUser().getEmail();
  var url = SERVER_URL + "/email/" + emailId;
  
  try {
    var options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'headers': {
        'user-email': userEmail
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();
    
    if (statusCode === 200) {
      return JSON.parse(contentText);
    } else {
      console.error('Error fetching email. Status code:', statusCode, 'Response:', contentText);
      return null;
    }
  } catch (error) {
    console.error('Error in fetchSavedEmailById:', error);
    return null;
  }
}


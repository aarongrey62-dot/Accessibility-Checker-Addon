/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  SlidesApp.getUi()
      .createMenu('Accessibility Checker')
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Accessibility Checker') 
      .setWidth(300);
  SlidesApp.getUi().showSidebar(html);
}

function parseColor(color) {
  if (!color) return "None/Transparent";
  var type = color.getColorType();
  if (type === SlidesApp.ColorType.RGB) {
    var rgb = color.asRgbColor();
    var r = rgb.getRed(), g = rgb.getGreen(), b = rgb.getBlue();
    return "#" + (1 << 24 | r << 16 | g << 8 | b).toString(16).slice(1).toUpperCase();
  } else if (type === SlidesApp.ColorType.THEME) {
    return "Theme: " + color.asThemeColor().getThemeColorType().toString();
  }
  return "Unknown";
}

function getCurrentSlideData() {
  var selection = SlidesApp.getActivePresentation().getSelection();
  var slide = selection.getCurrentPage();
  if (!slide) return { error: "Select a slide to audit." };

  var allSlides = SlidesApp.getActivePresentation().getSlides();
  var slideNum = 0;
  for (var i = 0; i < allSlides.length; i++) {
    if (allSlides[i].getObjectId() === slide.getObjectId()) { slideNum = i + 1; break; }
  }

  var elements = slide.getPageElements();
  var data = [];
  var slideTitleText = null; 

  elements.forEach(function(el, index) {
    var typeStr = el.getPageElementType().toString();
    var alt = el.getDescription() || "";
    var isExplicitlyDecorative = (alt.trim().toLowerCase() === "decorative"); 
    var displayName = typeStr; 
    var isTitle = false; 
    var textContent = "";

    if (typeStr === "SHAPE") {
      var shape = el.asShape();
      try { textContent = shape.getText().asString().trim().replace(/\n/g, " "); } catch(e) {}
      try {
        var pType = shape.getPlaceholderType();
        if (pType === SlidesApp.PlaceholderType.TITLE || pType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
          isTitle = true;
          slideTitleText = textContent || "[Empty Title]";
        }
      } catch(e) {}
      
      if (isExplicitlyDecorative) { displayName = 'Decorative Shape'; } 
      else if (textContent.length > 0) { displayName = 'Text: "' + textContent.substring(0, 30) + '..."'; } 
      else { displayName = alt ? 'Shape: "' + alt.substring(0, 25) + '..."' : 'Shape (No Alt-Text)'; }
    } else if (typeStr === "IMAGE") {
      displayName = isExplicitlyDecorative ? 'Decorative Image' : (alt ? 'Image: "' + alt.substring(0, 25) + '..."' : 'Image (No Alt-Text)');
    } 

    data.push({
      id: el.getObjectId(),
      type: typeStr,
      title: displayName,
      text: textContent || "", 
      order: index + 1,
      isTitle: isTitle,
      alt: alt,
      isDecorative: isExplicitlyDecorative
    });
  });

  return { elements: data, slideId: slide.getObjectId(), slideTitle: slideTitleText, slideNumber: slideNum, totalSlides: allSlides.length };
}

function saveNewOrder(orderedIds) {
  var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  if (!slide) return;
  var elements = slide.getPageElements();
  var elMap = {};
  elements.forEach(function(el) { elMap[el.getObjectId()] = el; });
  orderedIds.forEach(function(id) { if (elMap[id]) { elMap[id].bringToFront(); } });
}

function selectElement(objectId) {
  try {
    var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var elements = slide.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getObjectId() === objectId) { elements[i].select(); break; }
    }
  } catch (e) {}
}

function injectNativeTitle(suggestedText) {
  try {
    var presentation = SlidesApp.getActivePresentation();
    var currentSlide = presentation.getSelection().getCurrentPage();
    if (!currentSlide) throw new Error("No slide selected.");
    
    var tempSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    var titleShape = tempSlide.getPageElements().find(el => {
       try { 
         var pt = el.asShape().getPlaceholderType(); 
         return pt === SlidesApp.PlaceholderType.TITLE || pt === SlidesApp.PlaceholderType.CENTERED_TITLE; 
       } catch(e) { return false; }
    });
    
    if (titleShape) { 
      var newTitle = currentSlide.insertPageElement(titleShape);
      
      var finalTitle = (suggestedText && suggestedText.trim() !== "") ? suggestedText : "Enter Title Here";
      
      newTitle.asShape().getText().setText(finalTitle); 
      newTitle.sendToBack();
    }
    tempSlide.remove();
  } catch (e) { throw new Error(e.message); }
}

function getSlideImageBase64(slideObjectId) {
  var presentationId = SlidesApp.getActivePresentation().getId();
  var thumbnail = Slides.Presentations.Pages.getThumbnail(presentationId, slideObjectId, { "thumbnailProperties.thumbnailSize": "MEDIUM" });
  var options = { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  var response = UrlFetchApp.fetch(thumbnail.contentUrl, options);
  if (response.getResponseCode() !== 200) throw new Error("Failed to capture slide image.");
  var blob = response.getBlob();
  return { mimeType: blob.getContentType(), base64Data: Utilities.base64Encode(blob.getBytes()) };
}

function analyzeOrderAndTitle(payload) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('VERTEX_AI_KEY');
  if (!apiKey) return JSON.stringify({ error: "Missing API Key." });

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=' + apiKey;
  
  var systemInstructionText = `Analyze this slide's accessibility. 
1. Confidence score (0-100) for the current title relevance.
2. Suggest a concise title (maximum 3 words, ideally 1-2 words).
3. Logical reading order (top-to-bottom, left-to-right). 

Return ONLY a JSON object matching this exact structure:
{
  "title_confidence_score": 0,
  "suggested_title": "Example Title",
  "suggested_order_ids": ["id1", "id2"],
  "suggested_order_string": "1.[Title] > 2.[Text]..."
}`;
  
  var slideImage = null;
  try { slideImage = getSlideImageBase64(payload.slideId); } catch(e) {}

  var contentsParts = [{ "text": "Analyze this slide data: " + JSON.stringify(payload) }];
  if (slideImage) {
    contentsParts.push({ "inline_data": { "mime_type": slideImage.mimeType, "data": slideImage.base64Data } });
  }

  var apiPayload = { 
    "system_instruction": { "parts": [{ "text": systemInstructionText }] }, 
    "contents": [{ "parts": contentsParts }],
    // 3. JSON ENFORCER: Prevents markdown wrapping and silent crashes
    "generationConfig": { "response_mime_type": "application/json" } 
  };

  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(apiPayload), "muteHttpExceptions": true };
// NEW: Retry Loop (Tries up to 3 times before failing)
  var maxRetries = 3;
  var response = null;
  var errorMsg = "";

  for (var i = 0; i < maxRetries; i++) {
    try {
      response = UrlFetchApp.fetch(url, options);
      
      // If the API connects successfully (Code 200)
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.candidates && json.candidates[0].content.parts[0].text) {
          return json.candidates[0].content.parts[0].text;
        } else {
          return JSON.stringify({ error: "Empty AI response." });
        }
      } else {
        errorMsg = "API returned " + response.getResponseCode();
      }
    } catch(e) {
      errorMsg = e.toString();
    }
    
    // If it fails, sleep for a second or two, then the loop tries again!
    if (i < maxRetries - 1) Utilities.sleep(1000 * (i + 1)); 
  }

  // If it completely fails 3 times in a row, send the error back to the sidebar
  return JSON.stringify({ error: "API limit hit. " + errorMsg });
}

function getActiveSlideIdOnly() {
  try {
    var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    return slide ? slide.getObjectId() : null;
  } catch(e) {
    return null;
  }
}

function generateDecorativeBackground(slideId, decorativeIds) {
  if (!decorativeIds || decorativeIds.length === 0) return { error: "No elements selected." };

  try {
    var presentation = SlidesApp.getActivePresentation();
    var presentationId = presentation.getId();
    var originalSlide = presentation.getSlideById(slideId);

    // STEP 1: Duplicate the original slide
    var tempSlide = originalSlide.duplicate();
    var tempSlideId = tempSlide.getObjectId();

    // STEP 2: Remove all readable elements from the temp slide
    var tempElements = tempSlide.getPageElements();
    var originalElements = originalSlide.getPageElements();
    
    for (var i = originalElements.length - 1; i >= 0; i--) {
       var origId = originalElements[i].getObjectId();
       if (decorativeIds.indexOf(origId) === -1) {
          tempElements[i].remove();
       }
    }

    // STEP 3: Force Google to save the file so the API can see the new temp slide
    presentation.saveAndClose(); 
    Utilities.sleep(1500); 

    // STEP 4: Take the snapshot
    var thumbnail = Slides.Presentations.Pages.getThumbnail(presentationId, tempSlideId, { "thumbnailProperties.thumbnailSize": "LARGE" });
    var options = { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    var response = UrlFetchApp.fetch(thumbnail.contentUrl, options);
    var blob = response.getBlob();
    
    var base64Data = Utilities.base64Encode(blob.getBytes());
    var mimeType = blob.getContentType();
    var dataUri = "data:" + mimeType + ";base64," + base64Data;

    // STOP HERE! Don't delete the slide yet. 
    // Return the image data AND the temp slide ID to the sidebar.
    return { success: true, imageData: dataUri, tempSlideId: tempSlideId };

  } catch(e) {
    return { error: e.toString() };
  }
}

// NEW FUNCTION: A dedicated cleaner that runs in a fresh execution so your screen instantly updates!
function cleanupTempSlide(tempSlideId) {
  try {
    var pres = SlidesApp.getActivePresentation();
    var slide = pres.getSlideById(tempSlideId);
    if (slide) slide.remove();
  } catch(e) {}
}

function applyFlattenedImage(slideId, dataUri, decorativeIds) {
  try {
    var presentation = SlidesApp.getActivePresentation();
    var slide = presentation.getSlideById(slideId);

    // Convert the Data URI back into an image file
    var base64Data = dataUri.substring(dataUri.indexOf(',') + 1);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', 'Flattened_Background.png');

    // Insert the image onto the slide
    var newBgImage = slide.insertImage(blob);

    // Stretch the image to perfectly cover the entire slide
    var pageWidth = presentation.getPageWidth();
    var pageHeight = presentation.getPageHeight();
    newBgImage.setLeft(0).setTop(0).setWidth(pageWidth).setHeight(pageHeight);

    // Send it to the back of Layer 3 (Above the master theme, behind your text)
    newBgImage.sendToBack();

    // Crucial Accessibility Step: Tag it as decorative so the screen reader ignores it
    newBgImage.setDescription("decorative");

    // Delete the original vector shapes
    var elements = slide.getPageElements();
    elements.forEach(function(el) {
      if (decorativeIds.indexOf(el.getObjectId()) !== -1) {
        el.remove();
      }
    });

    return { success: true };
  } catch(e) {
    return { error: e.toString() };
  }
}

function updateElementAltText(objectId, newText) {
  try {
    var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var elements = slide.getPageElements();
    
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].getObjectId() === objectId) {
        elements[i].setDescription(newText); // This sets the actual Alt-Text!
        return { success: true };
      }
    }
    return { error: "Element not found." };
  } catch (e) {
    return { error: e.toString() };
  }
}

function generateAltTextSuggestions(payload) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('VERTEX_AI_KEY');
  if (!apiKey) return JSON.stringify({ error: "Missing API Key." });

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=' + apiKey;
  
  var systemInstructionText = `Analyze this slide's visual layout and context. 
You will be given a JSON list of element IDs and types (IMAGE or SHAPE) that are currently missing alt-text.

For each IMAGE: Generate a highly succinct, descriptive alt-text summary. Be extremely concise (maximum 1 short sentence, under 15 words). Focus only on the core subject and ignore minor background details.
For each SHAPE: Evaluate if the shape provides essential contextual or informational meaning to the slide, or if it is purely a decorative/stylistic element (e.g., background waves, accent blocks, framing banners, colored rectangles). Provide a confidence score from 0 to 100 on how likely it is to be purely decorative, along with a brief 1-sentence reason.

Return ONLY a JSON object matching this exact structure:
{
  "element_id_here": {
    "type": "IMAGE",
    "suggestion": "London skyline featuring the London Eye and Big Ben."
  },
  "another_element_id": {
    "type": "SHAPE",
    "is_decorative_prob": 85,
    "reason": "It is a decorative wave graphic at the bottom providing purely visual styling."
  }
}`;
  
  var slideImage = null;
  try { slideImage = getSlideImageBase64(payload.slideId); } catch(e) {}

  var contentsParts = [{ "text": "Elements to evaluate: " + JSON.stringify(payload.elements) }];
  if (slideImage) {
    contentsParts.push({ "inline_data": { "mime_type": slideImage.mimeType, "data": slideImage.base64Data } });
  }

  var apiPayload = { 
    "system_instruction": { "parts": [{ "text": systemInstructionText }] }, 
    "contents": [{ "parts": contentsParts }],
    "generationConfig": { "response_mime_type": "application/json" } 
  };

  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(apiPayload), "muteHttpExceptions": true };

  var maxRetries = 3;
  for (var i = 0; i < maxRetries; i++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.candidates && json.candidates[0].content.parts[0].text) {
          return json.candidates[0].content.parts[0].text;
        }
      }
    } catch(e) { }
    if (i < maxRetries - 1) Utilities.sleep(1000 * (i + 1)); 
  }
  return JSON.stringify({ error: "API connection failed." });
}
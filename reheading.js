/**
 * @OnlyCurrentDoc
 *
 * 
 * Copyright 2024 app_support@g.irfu.se
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 * Known issues:
 * - If hiding headings from TOC links will not be updated to that level
 * - It would be nice to combine all manual button presses into one. However
 *   it is impossible until Google fixes either
 *      https://issuetracker.google.com/u/0/issues/36761940
 *   or
 *      https://issuetracker.google.com/u/0/issues/36758222
 *   since heading links can only be extracted from TOC, and TOC cannot be 
 *   auto-generated.
 */

/**
 * Creates the menu items in the add-on menu.
 *
 * @param {Event} e The onOpen event.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Update Headings', 'updateAllHeadings')
    .addItem('Update Links from TOC', 'replaceHeadingLinks')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * Update all headings in the document to use a numeric prefix.
 * The order is as follows:
 * Heading 1: 1. 2. 3. ...
 * Heading 2: 1.1. 1.2. 1.3. ...
 * Heading 3: 1.1.1. 1.1.2. 1.1.3. ...
 * ...
 * This function will skip over any heading that appears after an Annex or Appendix section.
 * 
 * This function will not change any headings that already have the correct prefix.
 */
function updateAllHeadings() {
  var curDoc = DocumentApp.getActiveDocument();
  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchHeading = DocumentApp.ParagraphHeading;
  var searchResult = null;
  var h1 = 0;
  var h2 = 0;
  var h3 = 0;
  var h4 = 0;
  var h5 = 0;
  var h6 = 0;

  while (searchResult = curDoc.getBody().findElement(searchType, searchResult)) {
    var par = searchResult.getElement().asParagraph();


    if (par.getHeading() == searchHeading.HEADING1) {
      if (par.getText().match(/(^A\w*[xX] [0-9A-Z]:)/g))
        return; //if Annex/Appendix section stop renumbering
      h1++;
      h2 = 0;
      h3 = 0;
      h4 = 0;
      h5 = 0;
      h6 = 0;
      replaceHeading(par, h1 + '. ');
    } else if (par.getHeading() == searchHeading.HEADING2) {
      h2++;
      h3 = 0;
      h4 = 0;
      h5 = 0;
      h6 = 0;
      replaceHeading(par, h1 + '.' + h2 + '. ');
    } else if (par.getHeading() == searchHeading.HEADING3) {
      h3++;
      h4 = 0;
      h5 = 0;
      h6 = 0;
      replaceHeading(par, h1 + '.' + h2 + '.' + h3 + '. ');
    } else if (par.getHeading() == searchHeading.HEADING4) {
      h4++;
      h5 = 0;
      h6 = 0;
      replaceHeading(par, h1 + '.' + h2 + '.' + h3 + '.' + h4 + '. ');
    } else if (par.getHeading() == searchHeading.HEADING5) {
      h5++;
      h6 = 0;
      replaceHeading(par, h1 + '.' + h2 + '.' + h3 + '.' + h4 + '.' + h5 + '. ');
    } else if (par.getHeading() == searchHeading.HEADING6) {
      h6++;
      replaceHeading(par, h1 + '.' + h2 + '.' + h3 + '.' + h4 + '.' + h5 + '.' + h6 + '. ');
    }
  }
}

/**
 * Replace the heading prefix with the given new prefix.
 * If the new prefix is the same as the old one, do nothing.
 * If the heading does not have a prefix, insert the new prefix
 * at the beginning of the heading.
 *
 * @param {DocumentApp.Element} heading The heading element to modify.
 * @param {String} newPrefix The new prefix to use.
 */
function replaceHeading(heading, newPrefix) {
  let oldPrefixes = heading.getText().match(/(^[0-9\.]+ )/g);
  Logger.log(heading.getText())
  if (oldPrefixes) {
    if (oldPrefixes[0] == newPrefix) {
      return;
    }
    heading.replaceText('^[0-9+\.]+ ', newPrefix);
  } else {
    heading.insertText(0, newPrefix);
  }
}

/**
 * Updates all links in the document that point to headings with the
 * new heading text. This is necessary because the heading text may
 * have changed after the user clicked on "Update Headings".
 *
 * The function goes through all links in the document, checks if the
 * link points to a heading and if so, updates the link to point to
 * the new heading text. If the link does not point to a heading
 * anymore, it is not updated.
 *
 * The function logs all links that could not be updated.
 * 
 * Credits: https://stackoverflow.com/questions/55923420/update-link-to-heading-in-google-docs
 */
function replaceHeadingLinks() {
  var curDoc = DocumentApp.getActiveDocument();
  var links = getAllLinks_(curDoc.getBody());
  var headings = getAllHeadings_(curDoc.getBody());
  var deprecatedLinks = []; // holds all links to headings that do not exist anymore.

  links.forEach(function (link) {

    if (link.url.startsWith('#heading')) {

      // get the new heading text
      var newHeadingText = headings.get(link.url);

      // if the link does not exist anymore, we cannot update it.
      if (typeof newHeadingText !== "undefined") {
        var sectionNumber = newHeadingText.match(/^[0-9\.]+ /g);
        var annexText = newHeadingText.match(/^A\w*[xX] [0-9A-Z]: /g);
        if (sectionNumber)
          newHeadingText = 'Section ' + sectionNumber[0].slice(0, -1);
        else if (annexText)
          newHeadingText = annexText[0].slice(0, -2);
        var newOffset = link.startOffset + newHeadingText.length - 1;

        // delete the old text, insert new one and set link
        link.element.deleteText(link.startOffset, link.endOffsetInclusive);
        link.element.insertText(link.startOffset, newHeadingText);
        link.element.setLinkUrl(link.startOffset, newOffset, link.url);

      } else {
        deprecatedLinks.push(link);
      }

    }

  }
  )

  // error handling: show deprecated links:

  if (deprecatedLinks.length > 0) {
    Logger.log("Links we could not update:");
    var list = "";
    for (var i = 0; i < deprecatedLinks.length; i++) {
      var link = deprecatedLinks[i];
      var oldText = link.element.getText().substring(link.startOffset, link.endOffsetInclusive);
      Logger.log("heading: " + link.url + " / description: " + oldText);
      list += "heading: " + link.url + " / description: " + oldText + "\n";
    }
    var ui = DocumentApp.getUi();
    ui.alert("Links we could not update:", list, ui.ButtonSet.OK);
  } else {
    Logger.log("all links updated");
  }

}


/**
 * Get an array of all LinkUrls in the document. The function is
 * recursive, and if no element is provided, it will default to
 * the active document's Body element.
 *
 * @param {Element} element The document element to operate on. 
 * .
 * @returns {Array}         Array of objects, vis
 *                              {element,
 *                               startOffset,
 *                               endOffsetInclusive, 
 *                               url}
 *
 * Credits: https://stackoverflow.com/questions/18727341/get-all-links-in-a-document/40730088
 */
function getAllLinks_(element) {
  var links = [];
  element = element || DocumentApp.getActiveDocument().getBody();

  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textObj = element.editAsText();
    var text = element.getText();
    var inUrl = false;
    var curUrl = {};
    for (var ch = 0; ch < text.length; ch++) {
      var url = textObj.getLinkUrl(ch);
      if (url != null) {
        if (!inUrl) {
          // We are now!
          inUrl = true;
          curUrl = {};
          curUrl.element = element;
          curUrl.url = String(url); // grab a copy
          curUrl.startOffset = ch;
        }
        else {
          curUrl.endOffsetInclusive = ch;
        }
      }
      else {
        if (inUrl) {
          // Not any more, we're not.
          inUrl = false;
          links.push(curUrl);  // add to links
          curUrl = {};
        }
      }
    }
    // edge case: link is at the end of a paragraph
    // check if object is empty
    if (inUrl && (Object.keys(curUrl).length !== 0 || curUrl.constructor !== Object)) {
      links.push(curUrl);  // add to links
      curUrl = {};
    }
  }
  else {
    // only traverse if the element is traversable
    if (typeof element.getNumChildren !== "undefined") {
      var numChildren = element.getNumChildren();

      for (var i = 0; i < numChildren; i++) {

        // exclude Table of Contents

        child = element.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.TABLE_OF_CONTENTS) {
          links = links.concat(getAllLinks_(element.getChild(i)));
        }
      }
    }
  }

  return links;
}


/**
 * returns a map of all headings within an element. The map key
 * is the heading ID, such as h.q1xuchg2smrk
 *
 * THIS REQUIRES A CURRENT TABLE OF CONTENTS IN THE DOCUMENT TO WORK PROPERLY.
 *
 * @param {Element} element The document element to operate on. 
 * .
 * @returns {Map} Map with heading ID as key and the heading element as value.
 */
function getAllHeadings_(element) {

  var headingsMap = new Map();
  var realHeadings = getRealHeadings_(element);
  var isUpToDate = true;
  var p = element.findElement(DocumentApp.ElementType.TABLE_OF_CONTENTS).getElement();

  if (p !== null) {
    var toc = p.asTableOfContents();
    for (var ti = 0; ti < toc.getNumChildren(); ti++) {

      var itemToc = toc.getChild(ti).asParagraph().getChild(0).asText();
      var itemText = itemToc.getText();
      var itemUrl = itemToc.getLinkUrl(0);
      var itemDesc = null;

      // strip the line numbers if TOC contains line numbers
      var itemText = itemText.match(/(.*)\t/)[1];
      if (itemText.trim() != realHeadings[ti]) {
        isUpToDate = false;
      }
      headingsMap.set(itemUrl, itemText);
    }
  }

  if (!isUpToDate) {
    var ui = DocumentApp.getUi();
    ui.alert("Table of Contents out of date", "Please update the Table of Contents in the document.\n(This can also be the result of hidden heading levels.)", ui.ButtonSet.OK);
  }
  return headingsMap;
}

/**
 * Gets all headings from an element.
 *
 * @param {Element} element The document element to search within.
 * @returns {Array<String>} Array of strings, each being the text of a heading in the document.
 */
function getRealHeadings_(element) {

  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchHeading = DocumentApp.ParagraphHeading;
  var searchResult = null;
  var headings = [];
  while (searchResult = element.findElement(searchType, searchResult)) {
    var par = searchResult.getElement().asParagraph();
    if (par.getHeading() == searchHeading.HEADING1 || par.getHeading() == searchHeading.HEADING2 || par.getHeading() == searchHeading.HEADING3
      || par.getHeading() == searchHeading.HEADING4 || par.getHeading() == searchHeading.HEADING5 || par.getHeading() == searchHeading.HEADING6) {
      headings.push(par.getText().trim());
    }
  }
  return headings;

}

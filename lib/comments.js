const AdmZip = require('adm-zip');
const path = require('path');

/**
 * Extract comments from a PPTX file.
 * Returns an object: { slideNumber: [{ author, text, date }], ... }
 * Returns null if the file has no comments or is not a .pptx.
 */
function extractComments(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== '.pptx') return null;

  try {
    const zip = new AdmZip(filePath);
    const entries = zip.getEntries();

    // Find comment authors
    const authors = {};
    const authorsEntry = entries.find(function (e) {
      return e.entryName === 'ppt/commentAuthors.xml';
    });
    if (authorsEntry) {
      const xml = authorsEntry.getData().toString('utf8');
      // Parse <p:cmAuthor id="0" name="John" .../>
      const authorRegex = /<p:cmAuthor[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"[^>]*\/?>/g;
      var m;
      while ((m = authorRegex.exec(xml)) !== null) {
        authors[m[1]] = m[2];
      }
    }

    // Build slide number mapping from relationships
    // ppt/slides/slide1.xml -> slide 1, etc.
    // Comments reference slides via ppt/comments/commentN.xml
    // The link is in ppt/slides/_rels/slideN.xml.rels

    // First, find which comment files map to which slides
    const slideCommentMap = {}; // commentFileName -> slideNumber

    for (var i = 0; i < entries.length; i++) {
      var entryName = entries[i].entryName;
      // Match ppt/slides/_rels/slide{N}.xml.rels
      var relsMatch = entryName.match(/^ppt\/slides\/_rels\/slide(\d+)\.xml\.rels$/);
      if (relsMatch) {
        var slideNum = parseInt(relsMatch[1]);
        var relsXml = entries[i].getData().toString('utf8');
        // Find references to ../comments/commentN.xml
        var commentRefRegex = /Target="\.\.\/comments\/(comment\d+\.xml)"/g;
        var rm;
        while ((rm = commentRefRegex.exec(relsXml)) !== null) {
          slideCommentMap[rm[1]] = slideNum;
        }
      }
    }

    // Now parse each comment file
    var result = {};
    var hasAny = false;

    for (var j = 0; j < entries.length; j++) {
      var eName = entries[j].entryName;
      var cmMatch = eName.match(/^ppt\/comments\/(comment\d+\.xml)$/);
      if (!cmMatch) continue;

      var commentFileName = cmMatch[1];
      var slideNum2 = slideCommentMap[commentFileName];
      if (!slideNum2) continue;

      var cmXml = entries[j].getData().toString('utf8');

      // Parse <p:cm authorId="0" dt="2024-01-15T10:30:00"> ... <p:text>comment text</p:text> ... </p:cm>
      var cmRegex = /<p:cm\b[^>]*authorId="(\d+)"[^>]*(?:dt="([^"]*)")?[^>]*>([\s\S]*?)<\/p:cm>/g;
      var cm;
      while ((cm = cmRegex.exec(cmXml)) !== null) {
        var authorId = cm[1];
        var dt = cm[2] || '';
        var inner = cm[3];

        // Extract text
        var textMatch = inner.match(/<p:text>([\s\S]*?)<\/p:text>/);
        var text = textMatch ? textMatch[1].trim() : '';
        if (!text) continue;

        var author = authors[authorId] || 'Unknown';

        if (!result[slideNum2]) result[slideNum2] = [];
        result[slideNum2].push({
          author: author,
          text: text,
          date: dt,
        });
        hasAny = true;
      }
    }

    return hasAny ? result : null;
  } catch (e) {
    return null;
  }
}

module.exports = { extractComments };

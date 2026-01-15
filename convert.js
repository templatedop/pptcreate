const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const cheerio = require('cheerio');

// Color mapping
const colors = {
  blue: '1e40af',
  teal: '06b6d4',
  purple: '8b5cf6',
  orange: 'f97316',
  green: '10b981',
  indigo: '6366f1',
  pink: 'ec4899',
  yellow: 'eab308'
};

function extractBackgroundColor(html) {
  const $ = cheerio.load(html);
  const slideDiv = $('.slide');
  const style = slideDiv.attr('style') || '';

  // Check for gradient background
  if (style.includes('linear-gradient')) {
    return { type: 'gradient', value: style };
  }

  // Check for solid background
  const bgMatch = style.match(/background:\s*([^;]+)/);
  if (bgMatch) {
    return { type: 'solid', value: bgMatch[1] };
  }

  return { type: 'solid', value: '#f9fafb' };
}

function parseHTMLSlide(html, slideNumber) {
  const $ = cheerio.load(html);

  // Extract title
  const title = $('.title-zone h1').text().trim();

  // Extract background
  const background = extractBackgroundColor(html);

  // Extract content elements
  const elements = [];

  // Parse different content types
  $('.content-zone').find('*').each((i, elem) => {
    const $elem = $(elem);
    const text = $elem.text().trim();
    const tag = elem.name;

    if (text && !$elem.parent().is('h1, h2, h3, h4, p')) {
      elements.push({
        tag,
        text,
        classes: $elem.attr('class') || ''
      });
    }
  });

  return { title, background, elements, slideNumber };
}

async function convertSlides() {
  console.log('Starting APT 2.0 Presentation conversion...');

  // Create a new presentation
  const pptx = new PptxGenJS();

  // Set presentation properties
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'India Post';
  pptx.company = 'India Post';
  pptx.subject = 'APT 2.0 Presentation';
  pptx.title = 'APT 2.0 - Advanced Postal Technology';

  // Directory containing slides
  const slidesDir = path.join(__dirname, 'slides');

  // Get all HTML slide files and sort them
  const slideFiles = fs.readdirSync(slidesDir)
    .filter(file => file.endsWith('.html'))
    .sort();

  console.log(`Found ${slideFiles.length} slides to convert`);

  // Process each slide
  for (const slideFile of slideFiles) {
    const slideNumber = slideFile.match(/\d+/)[0];
    console.log(`Processing slide ${slideNumber}...`);

    const slidePath = path.join(slidesDir, slideFile);
    const htmlContent = fs.readFileSync(slidePath, 'utf8');

    // Parse HTML
    const slideData = parseHTMLSlide(htmlContent, slideNumber);

    // Add a new slide
    const slide = pptx.addSlide();

    try {
      // Set background based on slide type
      if (slideData.background.type === 'gradient') {
        // For gradient backgrounds (title slide, thank you slide)
        if (slideNumber === '01' || slideNumber === '21') {
          slide.background = { color: '1e40af' };
        }
      } else {
        slide.background = { color: 'f9fafb' };
      }

      // Add title if present
      if (slideData.title) {
        slide.addText(slideData.title, {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 0.75,
          fontSize: 34,
          bold: true,
          color: '1f2937',
          align: 'center'
        });
      }

      console.log(`✓ Slide ${slideNumber} converted successfully`);
    } catch (error) {
      console.error(`✗ Error converting slide ${slideNumber}:`, error.message);
    }
  }

  // Generate the PowerPoint file
  const outputFile = 'APT_2.0_Presentation.pptx';
  console.log(`\nGenerating ${outputFile}...`);

  await pptx.writeFile({ fileName: outputFile });

  console.log(`\n✓ Presentation created successfully: ${outputFile}`);
  console.log(`Total slides: ${slideFiles.length}`);
}

// Run the conversion
convertSlides().catch(error => {
  console.error('Conversion failed:', error);
  process.exit(1);
});

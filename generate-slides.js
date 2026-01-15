const pptxgen = require('pptxgenjs');

// Create presentation
const pptx = new pptxgen();

// Configure presentation
pptx.layout = 'LAYOUT_16x9';
pptx.author = 'Mail Booking System';
pptx.title = 'Mail Booking Features';

// Color scheme
const colors = {
  primary: '1e40af',      // Deep Blue
  secondary: '06b6d4',    // Teal
  accent: '8b5cf6',       // Purple
  success: '10b981',      // Green
  warning: 'f97316',      // Orange
  text: '1f2937',         // Dark Gray
  lightGray: 'f3f4f6',    // Light Gray
  white: 'FFFFFF'
};

// ==========================================
// SLIDE 1: Title Slide
// ==========================================
const slide1 = pptx.addSlide();

// Background gradient
slide1.background = {
  fill: '1e40af'
};

// Main title
slide1.addText('Mail Booking Features', {
  x: 1,
  y: 2.0,
  w: 8,
  h: 1.2,
  fontSize: 54,
  bold: true,
  color: colors.white,
  align: 'center'
});

// Subtitle
slide1.addText('Modern, Efficient & User-Friendly', {
  x: 1,
  y: 3.3,
  w: 8,
  h: 0.5,
  fontSize: 24,
  color: colors.white,
  align: 'center',
  transparency: 20
});

// Decorative line
slide1.addShape(pptx.ShapeType.rect, {
  x: 3.5,
  y: 4.0,
  w: 3,
  h: 0.05,
  fill: { color: colors.secondary }
});

// ==========================================
// SLIDE 2: Mail Booking Features - Card Layout
// ==========================================
const slide2 = pptx.addSlide();

// Background
slide2.background = { fill: colors.lightGray };

// Title
slide2.addText('Mail Booking Features', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary,
  align: 'left'
});

// Subtitle
slide2.addText('Comprehensive booking system for mail services', {
  x: 0.5,
  y: 1.0,
  w: 9,
  h: 0.3,
  fontSize: 16,
  color: colors.text,
  align: 'left',
  transparency: 30
});

// Feature cards data
const features = [
  {
    icon: '📅',
    title: 'Easy Scheduling',
    description: 'Book mail pickups and deliveries with flexible time slots',
    color: colors.primary
  },
  {
    icon: '📍',
    title: 'Location Tracking',
    description: 'Real-time tracking of your mail items from pickup to delivery',
    color: colors.secondary
  },
  {
    icon: '💳',
    title: 'Multiple Payment Options',
    description: 'Support for credit cards, digital wallets, and cash on delivery',
    color: colors.accent
  },
  {
    icon: '🔔',
    title: 'Smart Notifications',
    description: 'Instant alerts via SMS, email, and push notifications',
    color: colors.success
  },
  {
    icon: '📊',
    title: 'Booking History',
    description: 'Complete record of all bookings with easy access to details',
    color: colors.warning
  },
  {
    icon: '🔒',
    title: 'Secure Platform',
    description: 'End-to-end encryption ensuring data privacy and security',
    color: colors.primary
  }
];

// Create 2x3 grid of feature cards
const cardWidth = 2.8;
const cardHeight = 1.4;
const cardMargin = 0.25;
const startX = 0.5;
const startY = 1.7;

features.forEach((feature, index) => {
  const row = Math.floor(index / 3);
  const col = index % 3;
  const x = startX + (col * (cardWidth + cardMargin));
  const y = startY + (row * (cardHeight + cardMargin));

  // Card background
  slide2.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: cardWidth,
    h: cardHeight,
    fill: { color: colors.white },
    line: { color: feature.color, width: 0.5, transparency: 50 },
    rectRadius: 0.1
  });

  // Accent bar on left
  slide2.addShape(pptx.ShapeType.rect, {
    x: x,
    y: y,
    w: 0.08,
    h: cardHeight,
    fill: { color: feature.color }
  });

  // Icon
  slide2.addText(feature.icon, {
    x: x + 0.2,
    y: y + 0.15,
    w: 0.5,
    h: 0.5,
    fontSize: 32,
    align: 'center'
  });

  // Title
  slide2.addText(feature.title, {
    x: x + 0.15,
    y: y + 0.6,
    w: cardWidth - 0.3,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: colors.text,
    align: 'left'
  });

  // Description
  slide2.addText(feature.description, {
    x: x + 0.15,
    y: y + 0.92,
    w: cardWidth - 0.3,
    h: 0.4,
    fontSize: 10,
    color: colors.text,
    align: 'left',
    transparency: 30
  });
});

// ==========================================
// SLIDE 3: Key Benefits
// ==========================================
const slide3 = pptx.addSlide();

slide3.background = { fill: colors.white };

// Title
slide3.addText('Key Benefits', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

// Benefits data
const benefits = [
  {
    number: '01',
    title: 'Time Saving',
    description: 'Reduce booking time by 70% with our streamlined process',
    color: colors.primary
  },
  {
    number: '02',
    title: 'Cost Effective',
    description: 'Competitive pricing with transparent fee structure',
    color: colors.secondary
  },
  {
    number: '03',
    title: 'User Friendly',
    description: 'Intuitive interface designed for all age groups',
    color: colors.accent
  },
  {
    number: '04',
    title: '24/7 Availability',
    description: 'Book anytime, anywhere with mobile and web access',
    color: colors.success
  }
];

const benefitStartY = 1.5;
const benefitHeight = 0.95;
const benefitSpacing = 0.15;

benefits.forEach((benefit, index) => {
  const y = benefitStartY + (index * (benefitHeight + benefitSpacing));

  // Number circle
  slide3.addShape(pptx.ShapeType.ellipse, {
    x: 0.8,
    y: y + 0.15,
    w: 0.6,
    h: 0.6,
    fill: { color: benefit.color }
  });

  slide3.addText(benefit.number, {
    x: 0.8,
    y: y + 0.15,
    w: 0.6,
    h: 0.6,
    fontSize: 20,
    bold: true,
    color: colors.white,
    align: 'center',
    valign: 'middle'
  });

  // Background bar
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: 1.6,
    y: y,
    w: 7.4,
    h: benefitHeight,
    fill: { color: colors.lightGray },
    rectRadius: 0.1
  });

  // Title
  slide3.addText(benefit.title, {
    x: 1.8,
    y: y + 0.15,
    w: 7,
    h: 0.3,
    fontSize: 18,
    bold: true,
    color: colors.text
  });

  // Description
  slide3.addText(benefit.description, {
    x: 1.8,
    y: y + 0.5,
    w: 7,
    h: 0.35,
    fontSize: 13,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 4: Booking Process Flow
// ==========================================
const slide4 = pptx.addSlide();

slide4.background = {
  fill: colors.white
};

// Title
slide4.addText('Simple Booking Process', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

// Process steps
const steps = [
  { step: '1', title: 'Select Service', icon: '📦' },
  { step: '2', title: 'Choose Date & Time', icon: '📅' },
  { step: '3', title: 'Enter Details', icon: '✍️' },
  { step: '4', title: 'Make Payment', icon: '💳' },
  { step: '5', title: 'Confirmation', icon: '✅' }
];

const stepWidth = 1.5;
const stepStartX = 0.7;
const stepY = 2.5;
const arrowY = stepY + 0.4;

steps.forEach((step, index) => {
  const x = stepStartX + (index * 1.8);

  // Step circle
  slide4.addShape(pptx.ShapeType.ellipse, {
    x: x,
    y: stepY - 0.5,
    w: 0.7,
    h: 0.7,
    fill: { color: colors.primary },
    line: { color: colors.primary, width: 2 }
  });

  // Step number
  slide4.addText(step.step, {
    x: x,
    y: stepY - 0.5,
    w: 0.7,
    h: 0.7,
    fontSize: 22,
    bold: true,
    color: colors.white,
    align: 'center',
    valign: 'middle'
  });

  // Icon
  slide4.addText(step.icon, {
    x: x - 0.1,
    y: stepY + 0.4,
    w: 0.9,
    h: 0.6,
    fontSize: 40,
    align: 'center'
  });

  // Title
  slide4.addText(step.title, {
    x: x - 0.3,
    y: stepY + 1.1,
    w: stepWidth,
    h: 0.4,
    fontSize: 12,
    bold: true,
    color: colors.text,
    align: 'center'
  });

  // Arrow (except for last step)
  if (index < steps.length - 1) {
    slide4.addShape(pptx.ShapeType.rightArrow, {
      x: x + 0.8,
      y: arrowY,
      w: 0.8,
      h: 0.25,
      fill: { color: colors.secondary }
    });
  }
});

// ==========================================
// SLIDE 5: Statistics & Impact
// ==========================================
const slide5 = pptx.addSlide();

// Gradient background
slide5.background = {
  fill: colors.primary
};

// Title
slide5.addText('Platform Statistics', {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.white
});

// Stats
const stats = [
  { number: '50K+', label: 'Active Users', icon: '👥' },
  { number: '200K+', label: 'Bookings Completed', icon: '📦' },
  { number: '95%', label: 'Customer Satisfaction', icon: '⭐' },
  { number: '24/7', label: 'Support Available', icon: '🔧' }
];

const statWidth = 2;
const statStartX = 0.75;
const statY = 2;

stats.forEach((stat, index) => {
  const x = statStartX + (index * 2.3);

  // Card
  slide5.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: statY,
    w: statWidth,
    h: 2,
    fill: { color: colors.white, transparency: 10 },
    line: { color: colors.white, width: 1, transparency: 30 },
    rectRadius: 0.15
  });

  // Icon
  slide5.addText(stat.icon, {
    x: x,
    y: statY + 0.3,
    w: statWidth,
    h: 0.5,
    fontSize: 40,
    align: 'center'
  });

  // Number
  slide5.addText(stat.number, {
    x: x,
    y: statY + 0.9,
    w: statWidth,
    h: 0.6,
    fontSize: 36,
    bold: true,
    color: colors.white,
    align: 'center'
  });

  // Label
  slide5.addText(stat.label, {
    x: x,
    y: statY + 1.5,
    w: statWidth,
    h: 0.4,
    fontSize: 13,
    color: colors.white,
    align: 'center',
    transparency: 20
  });
});

// ==========================================
// SLIDE 6: Thank You
// ==========================================
const slide6 = pptx.addSlide();

slide6.background = {
  fill: colors.primary
};

// Main text
slide6.addText('Thank You', {
  x: 1,
  y: 2.2,
  w: 8,
  h: 0.8,
  fontSize: 48,
  bold: true,
  color: colors.white,
  align: 'center'
});

// Subtitle
slide6.addText('Questions & Feedback Welcome', {
  x: 1,
  y: 3.1,
  w: 8,
  h: 0.4,
  fontSize: 20,
  color: colors.white,
  align: 'center',
  transparency: 20
});

// Save presentation
pptx.writeFile({ fileName: 'Mail_Booking_Features.pptx' })
  .then(() => {
    console.log('✅ Presentation created successfully!');
    console.log('📄 File: Mail_Booking_Features.pptx');
    console.log('📊 Slides: 6');
    console.log('');
    console.log('Slide Overview:');
    console.log('  1. Title Slide - Main introduction');
    console.log('  2. Mail Booking Features - 6 feature cards with icons');
    console.log('  3. Key Benefits - 4 numbered benefits');
    console.log('  4. Booking Process - 5-step flow diagram');
    console.log('  5. Statistics - Platform metrics and achievements');
    console.log('  6. Thank You - Closing slide');
  })
  .catch(err => {
    console.error('❌ Error creating presentation:', err);
  });

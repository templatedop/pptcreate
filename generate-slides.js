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
// SLIDE 6: User Interface Overview
// ==========================================
const slide6 = pptx.addSlide();

slide6.background = { fill: colors.lightGray };

slide6.addText('User Interface Overview', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

slide6.addText('Intuitive design across all platforms', {
  x: 0.5,
  y: 1.0,
  w: 9,
  h: 0.3,
  fontSize: 16,
  color: colors.text,
  transparency: 30
});

const uiFeatures = [
  { icon: '📱', title: 'Responsive Design', desc: 'Seamless experience on mobile, tablet, and desktop' },
  { icon: '🎨', title: 'Modern UI/UX', desc: 'Clean, intuitive interface following latest design trends' },
  { icon: '♿', title: 'Accessibility', desc: 'WCAG compliant for users with disabilities' },
  { icon: '🌐', title: 'Multi-language', desc: 'Support for 15+ languages and regional formats' }
];

uiFeatures.forEach((feature, index) => {
  const y = 1.7 + (index * 0.95);

  slide6.addShape(pptx.ShapeType.roundRect, {
    x: 1.5,
    y: y,
    w: 7,
    h: 0.85,
    fill: { color: colors.white },
    rectRadius: 0.1
  });

  slide6.addText(feature.icon, {
    x: 1.7,
    y: y + 0.15,
    w: 0.5,
    h: 0.5,
    fontSize: 32
  });

  slide6.addText(feature.title, {
    x: 2.4,
    y: y + 0.15,
    w: 6,
    h: 0.25,
    fontSize: 16,
    bold: true,
    color: colors.text
  });

  slide6.addText(feature.desc, {
    x: 2.4,
    y: y + 0.45,
    w: 5.8,
    h: 0.3,
    fontSize: 12,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 7: Mobile App Features
// ==========================================
const slide7 = pptx.addSlide();

slide7.background = { fill: colors.white };

slide7.addText('Mobile App Features', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const mobileFeatures = [
  { icon: '📲', title: 'QR Code Scanning', desc: 'Quick booking via QR code', color: colors.primary },
  { icon: '🗺️', title: 'GPS Integration', desc: 'Auto-detect pickup locations', color: colors.secondary },
  { icon: '🔔', title: 'Push Notifications', desc: 'Real-time status updates', color: colors.accent },
  { icon: '📸', title: 'Photo Upload', desc: 'Attach package photos', color: colors.success },
  { icon: '💾', title: 'Offline Mode', desc: 'Draft bookings without internet', color: colors.warning },
  { icon: '👆', title: 'Biometric Auth', desc: 'Fingerprint & face recognition', color: colors.primary }
];

mobileFeatures.forEach((feature, index) => {
  const row = Math.floor(index / 3);
  const col = index % 3;
  const x = 0.5 + (col * 3.2);
  const y = 1.5 + (row * 1.7);

  slide7.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 2.9,
    h: 1.4,
    fill: { color: colors.lightGray },
    rectRadius: 0.15
  });

  slide7.addText(feature.icon, {
    x: x + 0.3,
    y: y + 0.2,
    w: 0.6,
    h: 0.6,
    fontSize: 36
  });

  slide7.addText(feature.title, {
    x: x + 0.2,
    y: y + 0.75,
    w: 2.5,
    h: 0.25,
    fontSize: 13,
    bold: true,
    color: colors.text
  });

  slide7.addText(feature.desc, {
    x: x + 0.2,
    y: y + 1.0,
    w: 2.5,
    h: 0.25,
    fontSize: 10,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 8: Web Portal Features
// ==========================================
const slide8 = pptx.addSlide();

slide8.background = { fill: colors.primary };

slide8.addText('Web Portal Features', {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.white
});

const webFeatures = [
  'Bulk Booking Management - Upload CSV for multiple bookings',
  'Advanced Reporting - Download detailed analytics and reports',
  'Address Book - Save frequently used addresses',
  'Scheduled Bookings - Plan future pickups in advance',
  'Invoice Management - Generate and download invoices',
  'Team Collaboration - Share bookings with team members'
];

webFeatures.forEach((feature, index) => {
  const y = 1.5 + (index * 0.65);

  slide8.addShape(pptx.ShapeType.rect, {
    x: 1,
    y: y,
    w: 0.15,
    h: 0.5,
    fill: { color: colors.secondary }
  });

  slide8.addText(feature, {
    x: 1.3,
    y: y + 0.05,
    w: 7.5,
    h: 0.4,
    fontSize: 14,
    color: colors.white,
    valign: 'middle'
  });
});

// ==========================================
// SLIDE 9: Admin Dashboard
// ==========================================
const slide9 = pptx.addSlide();

slide9.background = { fill: colors.lightGray };

slide9.addText('Admin Dashboard', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

slide9.addText('Comprehensive management and analytics', {
  x: 0.5,
  y: 1.0,
  w: 9,
  h: 0.3,
  fontSize: 16,
  color: colors.text,
  transparency: 30
});

const adminFeatures = [
  { icon: '📊', title: 'Real-time Analytics', color: colors.primary },
  { icon: '👥', title: 'User Management', color: colors.secondary },
  { icon: '💰', title: 'Revenue Tracking', color: colors.success },
  { icon: '⚙️', title: 'System Configuration', color: colors.accent },
  { icon: '📈', title: 'Performance Metrics', color: colors.warning },
  { icon: '🔍', title: 'Audit Logs', color: colors.primary }
];

adminFeatures.forEach((feature, index) => {
  const col = index % 3;
  const row = Math.floor(index / 3);
  const x = 0.5 + (col * 3.2);
  const y = 1.8 + (row * 1.1);

  slide9.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 2.9,
    h: 0.9,
    fill: { color: colors.white },
    line: { color: feature.color, width: 1 },
    rectRadius: 0.1
  });

  slide9.addText(feature.icon, {
    x: x + 0.3,
    y: y + 0.2,
    w: 0.5,
    h: 0.5,
    fontSize: 28
  });

  slide9.addText(feature.title, {
    x: x + 1.0,
    y: y + 0.28,
    w: 1.7,
    h: 0.35,
    fontSize: 13,
    bold: true,
    color: colors.text,
    valign: 'middle'
  });
});

// ==========================================
// SLIDE 10: Payment Gateway Integration
// ==========================================
const slide10 = pptx.addSlide();

slide10.background = { fill: colors.white };

slide10.addText('Payment Gateway Integration', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const paymentMethods = [
  { name: 'Credit/Debit Cards', icon: '💳', providers: 'Visa, Mastercard, Amex, Discover' },
  { name: 'Digital Wallets', icon: '📱', providers: 'Apple Pay, Google Pay, Samsung Pay' },
  { name: 'Net Banking', icon: '🏦', providers: 'All major banks supported' },
  { name: 'UPI', icon: '📲', providers: 'PhonePe, Paytm, GPay' },
  { name: 'Cash on Delivery', icon: '💵', providers: 'Pay when mail is delivered' }
];

paymentMethods.forEach((method, index) => {
  const y = 1.5 + (index * 0.75);

  slide10.addShape(pptx.ShapeType.roundRect, {
    x: 1,
    y: y,
    w: 8,
    h: 0.65,
    fill: { color: colors.lightGray },
    rectRadius: 0.1
  });

  slide10.addText(method.icon, {
    x: 1.2,
    y: y + 0.1,
    w: 0.45,
    h: 0.45,
    fontSize: 28
  });

  slide10.addText(method.name, {
    x: 1.8,
    y: y + 0.1,
    w: 2.5,
    h: 0.25,
    fontSize: 14,
    bold: true,
    color: colors.text
  });

  slide10.addText(method.providers, {
    x: 1.8,
    y: y + 0.37,
    w: 6.8,
    h: 0.2,
    fontSize: 11,
    color: colors.text,
    transparency: 40
  });
});

// ==========================================
// SLIDE 11: Notification System
// ==========================================
const slide11 = pptx.addSlide();

slide11.background = { fill: colors.secondary };

slide11.addText('Smart Notification System', {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.white
});

const notifications = [
  { type: 'SMS', desc: 'Text alerts for critical updates', icon: '📱' },
  { type: 'Email', desc: 'Detailed booking confirmations', icon: '📧' },
  { type: 'Push', desc: 'In-app real-time notifications', icon: '🔔' },
  { type: 'WhatsApp', desc: 'Status updates via WhatsApp', icon: '💬' }
];

notifications.forEach((notif, index) => {
  const col = index % 2;
  const row = Math.floor(index / 2);
  const x = 1 + (col * 4.3);
  const y = 1.8 + (row * 1.5);

  slide11.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 3.8,
    h: 1.2,
    fill: { color: colors.white, transparency: 15 },
    line: { color: colors.white, width: 1.5 },
    rectRadius: 0.15
  });

  slide11.addText(notif.icon, {
    x: x + 0.3,
    y: y + 0.2,
    w: 0.7,
    h: 0.7,
    fontSize: 40
  });

  slide11.addText(notif.type, {
    x: x + 1.2,
    y: y + 0.25,
    w: 2.3,
    h: 0.3,
    fontSize: 18,
    bold: true,
    color: colors.white
  });

  slide11.addText(notif.desc, {
    x: x + 1.2,
    y: y + 0.6,
    w: 2.3,
    h: 0.4,
    fontSize: 12,
    color: colors.white,
    transparency: 20
  });
});

// ==========================================
// SLIDE 12: Tracking & Monitoring
// ==========================================
const slide12 = pptx.addSlide();

slide12.background = { fill: colors.white };

slide12.addText('Tracking & Monitoring', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const trackingStages = [
  { stage: 'Booking\nConfirmed', icon: '✅', status: 'Completed' },
  { stage: 'Pickup\nScheduled', icon: '📅', status: 'Completed' },
  { stage: 'In Transit', icon: '🚚', status: 'Current' },
  { stage: 'Out for\nDelivery', icon: '📦', status: 'Pending' },
  { stage: 'Delivered', icon: '🎉', status: 'Pending' }
];

trackingStages.forEach((item, index) => {
  const x = 0.7 + (index * 1.8);
  const y = 2.5;

  const stageColor = item.status === 'Completed' ? colors.success :
                     item.status === 'Current' ? colors.warning :
                     colors.text;

  slide12.addShape(pptx.ShapeType.ellipse, {
    x: x,
    y: y - 0.5,
    w: 0.8,
    h: 0.8,
    fill: { color: stageColor },
    line: { color: stageColor, width: 2 }
  });

  slide12.addText(item.icon, {
    x: x,
    y: y - 0.5,
    w: 0.8,
    h: 0.8,
    fontSize: 28,
    align: 'center',
    valign: 'middle'
  });

  slide12.addText(item.stage, {
    x: x - 0.3,
    y: y + 0.5,
    w: 1.4,
    h: 0.6,
    fontSize: 11,
    bold: true,
    color: colors.text,
    align: 'center'
  });

  if (index < trackingStages.length - 1) {
    slide12.addShape(pptx.ShapeType.line, {
      x: x + 0.85,
      y: y - 0.1,
      w: 0.8,
      h: 0,
      line: { color: colors.text, width: 2, dashType: 'dash' }
    });
  }
});

// ==========================================
// SLIDE 13: Security Features
// ==========================================
const slide13 = pptx.addSlide();

slide13.background = { fill: colors.lightGray };

slide13.addText('Security Features', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const securityFeatures = [
  {
    icon: '🔐',
    title: 'End-to-End Encryption',
    desc: 'AES-256 encryption for all data transmission'
  },
  {
    icon: '🛡️',
    title: 'PCI DSS Compliant',
    desc: 'Secure payment processing standards'
  },
  {
    icon: '👁️',
    title: 'Two-Factor Authentication',
    desc: 'Additional security layer for account access'
  },
  {
    icon: '🔒',
    title: 'Data Privacy',
    desc: 'GDPR and CCPA compliant data handling'
  },
  {
    icon: '🚨',
    title: 'Fraud Detection',
    desc: 'AI-powered anomaly detection system'
  },
  {
    icon: '📝',
    title: 'Audit Trails',
    desc: 'Complete logging of all system activities'
  }
];

securityFeatures.forEach((feature, index) => {
  const col = index % 2;
  const row = Math.floor(index / 2);
  const x = 0.8 + (col * 4.5);
  const y = 1.5 + (row * 1.2);

  slide13.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 4,
    h: 1,
    fill: { color: colors.white },
    rectRadius: 0.1
  });

  slide13.addText(feature.icon, {
    x: x + 0.2,
    y: y + 0.15,
    w: 0.6,
    h: 0.6,
    fontSize: 32
  });

  slide13.addText(feature.title, {
    x: x + 1.0,
    y: y + 0.2,
    w: 2.8,
    h: 0.25,
    fontSize: 14,
    bold: true,
    color: colors.text
  });

  slide13.addText(feature.desc, {
    x: x + 1.0,
    y: y + 0.5,
    w: 2.8,
    h: 0.35,
    fontSize: 11,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 14: API Integration
// ==========================================
const slide14 = pptx.addSlide();

slide14.background = { fill: colors.white };

slide14.addText('API Integration', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

slide14.addText('Seamless integration with third-party services', {
  x: 0.5,
  y: 1.0,
  w: 9,
  h: 0.3,
  fontSize: 16,
  color: colors.text,
  transparency: 30
});

const apiFeatures = [
  { icon: '🔌', title: 'RESTful API', desc: 'Standard HTTP methods with JSON responses' },
  { icon: '📡', title: 'Webhooks', desc: 'Real-time event notifications' },
  { icon: '🔑', title: 'OAuth 2.0', desc: 'Secure authentication and authorization' },
  { icon: '📚', title: 'API Documentation', desc: 'Comprehensive Swagger/OpenAPI docs' },
  { icon: '⚡', title: 'Rate Limiting', desc: 'Fair usage policies and throttling' },
  { icon: '🧪', title: 'Sandbox Environment', desc: 'Test integration without affecting production' }
];

apiFeatures.forEach((feature, index) => {
  const col = index % 3;
  const row = Math.floor(index / 3);
  const x = 0.5 + (col * 3.2);
  const y = 1.8 + (row * 1.3);

  slide14.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 2.9,
    h: 1.1,
    fill: { color: colors.lightGray },
    rectRadius: 0.1
  });

  slide14.addText(feature.icon, {
    x: x + 0.3,
    y: y + 0.15,
    w: 0.5,
    h: 0.5,
    fontSize: 30
  });

  slide14.addText(feature.title, {
    x: x + 0.2,
    y: y + 0.6,
    w: 2.5,
    h: 0.2,
    fontSize: 13,
    bold: true,
    color: colors.text
  });

  slide14.addText(feature.desc, {
    x: x + 0.2,
    y: y + 0.82,
    w: 2.5,
    h: 0.22,
    fontSize: 10,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 15: Customer Support
// ==========================================
const slide15 = pptx.addSlide();

slide15.background = { fill: colors.accent };

slide15.addText('Customer Support', {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.white
});

slide15.addText('24/7 assistance through multiple channels', {
  x: 0.5,
  y: 1.1,
  w: 9,
  h: 0.3,
  fontSize: 18,
  color: colors.white,
  transparency: 20
});

const supportChannels = [
  { channel: 'Live Chat', availability: 'Instant responses', icon: '💬' },
  { channel: 'Phone Support', availability: '24/7 toll-free', icon: '📞' },
  { channel: 'Email Support', availability: '< 2 hour response', icon: '📧' },
  { channel: 'Help Center', availability: 'Self-service portal', icon: '❓' }
];

supportChannels.forEach((item, index) => {
  const col = index % 2;
  const row = Math.floor(index / 2);
  const x = 1.2 + (col * 4);
  const y = 2 + (row * 1.4);

  slide15.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 3.6,
    h: 1.1,
    fill: { color: colors.white, transparency: 15 },
    line: { color: colors.white, width: 1.5 },
    rectRadius: 0.15
  });

  slide15.addText(item.icon, {
    x: x + 0.3,
    y: y + 0.2,
    w: 0.6,
    h: 0.6,
    fontSize: 36
  });

  slide15.addText(item.channel, {
    x: x + 1.1,
    y: y + 0.25,
    w: 2.2,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: colors.white
  });

  slide15.addText(item.availability, {
    x: x + 1.1,
    y: y + 0.6,
    w: 2.2,
    h: 0.25,
    fontSize: 13,
    color: colors.white,
    transparency: 20
  });
});

// ==========================================
// SLIDE 16: Pricing Plans
// ==========================================
const slide16 = pptx.addSlide();

slide16.background = { fill: colors.lightGray };

slide16.addText('Pricing Plans', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const pricingPlans = [
  {
    name: 'Basic',
    price: 'Free',
    features: ['10 bookings/month', 'Email support', 'Basic tracking'],
    color: colors.text
  },
  {
    name: 'Pro',
    price: '$29/mo',
    features: ['Unlimited bookings', 'Priority support', 'Advanced analytics', 'API access'],
    color: colors.secondary,
    highlight: true
  },
  {
    name: 'Enterprise',
    price: 'Custom',
    features: ['Custom solutions', 'Dedicated account manager', 'SLA guarantee', 'On-premise option'],
    color: colors.primary
  }
];

pricingPlans.forEach((plan, index) => {
  const x = 0.8 + (index * 3.1);
  const y = 1.3;

  slide16.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 2.8,
    h: 3.5,
    fill: { color: plan.highlight ? plan.color : colors.white },
    line: { color: plan.color, width: plan.highlight ? 3 : 1 },
    rectRadius: 0.15
  });

  const textColor = plan.highlight ? colors.white : colors.text;

  slide16.addText(plan.name, {
    x: x,
    y: y + 0.3,
    w: 2.8,
    h: 0.35,
    fontSize: 20,
    bold: true,
    color: textColor,
    align: 'center'
  });

  slide16.addText(plan.price, {
    x: x,
    y: y + 0.75,
    w: 2.8,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: textColor,
    align: 'center'
  });

  plan.features.forEach((feature, fIndex) => {
    slide16.addText('✓ ' + feature, {
      x: x + 0.3,
      y: y + 1.5 + (fIndex * 0.4),
      w: 2.2,
      h: 0.3,
      fontSize: 11,
      color: textColor,
      align: 'left'
    });
  });
});

// ==========================================
// SLIDE 17: Implementation Timeline
// ==========================================
const slide17 = pptx.addSlide();

slide17.background = { fill: colors.white };

slide17.addText('Implementation Timeline', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const timeline = [
  { phase: 'Discovery', duration: 'Week 1-2', activities: 'Requirements gathering & planning' },
  { phase: 'Setup', duration: 'Week 3-4', activities: 'System configuration & integration' },
  { phase: 'Testing', duration: 'Week 5-6', activities: 'UAT and quality assurance' },
  { phase: 'Training', duration: 'Week 7', activities: 'User training & documentation' },
  { phase: 'Go Live', duration: 'Week 8', activities: 'Production deployment & support' }
];

timeline.forEach((item, index) => {
  const y = 1.5 + (index * 0.7);

  slide17.addShape(pptx.ShapeType.rect, {
    x: 0.8,
    y: y,
    w: 0.05,
    h: 0.6,
    fill: { color: colors.secondary }
  });

  slide17.addShape(pptx.ShapeType.ellipse, {
    x: 0.7,
    y: y + 0.2,
    w: 0.25,
    h: 0.25,
    fill: { color: colors.primary }
  });

  slide17.addText(item.phase, {
    x: 1.2,
    y: y,
    w: 2,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: colors.primary
  });

  slide17.addText(item.duration, {
    x: 1.2,
    y: y + 0.3,
    w: 2,
    h: 0.25,
    fontSize: 12,
    color: colors.secondary,
    bold: true
  });

  slide17.addText(item.activities, {
    x: 3.5,
    y: y + 0.15,
    w: 5,
    h: 0.3,
    fontSize: 13,
    color: colors.text,
    transparency: 30
  });
});

// ==========================================
// SLIDE 18: Success Stories
// ==========================================
const slide18 = pptx.addSlide();

slide18.background = { fill: colors.lightGray };

slide18.addText('Success Stories', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const successStories = [
  {
    company: 'City Logistics Co.',
    result: '300% increase in booking efficiency',
    quote: 'The platform transformed our operations completely'
  },
  {
    company: 'Express Couriers Ltd.',
    result: '50% reduction in operational costs',
    quote: 'Outstanding ROI within first 6 months'
  },
  {
    company: 'Metro Delivery Services',
    result: '95% customer satisfaction rate',
    quote: 'Our customers love the ease of booking'
  }
];

successStories.forEach((story, index) => {
  const y = 1.5 + (index * 1.2);

  slide18.addShape(pptx.ShapeType.roundRect, {
    x: 1,
    y: y,
    w: 8,
    h: 1,
    fill: { color: colors.white },
    rectRadius: 0.1
  });

  slide18.addText(story.company, {
    x: 1.3,
    y: y + 0.15,
    w: 7.4,
    h: 0.25,
    fontSize: 15,
    bold: true,
    color: colors.primary
  });

  slide18.addText(story.result, {
    x: 1.3,
    y: y + 0.43,
    w: 7.4,
    h: 0.2,
    fontSize: 13,
    color: colors.success,
    bold: true
  });

  slide18.addText('"' + story.quote + '"', {
    x: 1.3,
    y: y + 0.67,
    w: 7.4,
    h: 0.22,
    fontSize: 11,
    color: colors.text,
    italic: true,
    transparency: 30
  });
});

// ==========================================
// SLIDE 19: Technical Architecture
// ==========================================
const slide19 = pptx.addSlide();

slide19.background = { fill: colors.white };

slide19.addText('Technical Architecture', {
  x: 0.5,
  y: 0.4,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.primary
});

const techStack = [
  { layer: 'Frontend', tech: 'React, Vue.js, Mobile Apps (iOS/Android)', icon: '🎨', color: colors.secondary },
  { layer: 'Backend', tech: 'Node.js, Python, Microservices', icon: '⚙️', color: colors.accent },
  { layer: 'Database', tech: 'PostgreSQL, MongoDB, Redis Cache', icon: '💾', color: colors.success },
  { layer: 'Infrastructure', tech: 'AWS, Docker, Kubernetes', icon: '☁️', color: colors.warning },
  { layer: 'Security', tech: 'SSL/TLS, OAuth 2.0, Encryption', icon: '🔒', color: colors.primary }
];

techStack.forEach((item, index) => {
  const y = 1.5 + (index * 0.7);

  slide19.addShape(pptx.ShapeType.roundRect, {
    x: 1,
    y: y,
    w: 8,
    h: 0.6,
    fill: { color: item.color, transparency: 10 },
    line: { color: item.color, width: 1 },
    rectRadius: 0.08
  });

  slide19.addText(item.icon, {
    x: 1.2,
    y: y + 0.05,
    w: 0.5,
    h: 0.5,
    fontSize: 28
  });

  slide19.addText(item.layer, {
    x: 1.9,
    y: y + 0.12,
    w: 1.5,
    h: 0.35,
    fontSize: 15,
    bold: true,
    color: colors.text,
    valign: 'middle'
  });

  slide19.addText(item.tech, {
    x: 3.6,
    y: y + 0.12,
    w: 5,
    h: 0.35,
    fontSize: 13,
    color: colors.text,
    transparency: 30,
    valign: 'middle'
  });
});

// ==========================================
// SLIDE 20: Future Roadmap
// ==========================================
const slide20 = pptx.addSlide();

slide20.background = { fill: colors.primary };

slide20.addText('Future Roadmap', {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 0.6,
  fontSize: 36,
  bold: true,
  color: colors.white
});

const roadmap = [
  { quarter: 'Q2 2026', feature: 'AI-powered route optimization', icon: '🤖' },
  { quarter: 'Q3 2026', feature: 'Blockchain for package verification', icon: '⛓️' },
  { quarter: 'Q4 2026', feature: 'Drone delivery integration', icon: '🚁' },
  { quarter: 'Q1 2027', feature: 'AR package scanning', icon: '📱' }
];

roadmap.forEach((item, index) => {
  const col = index % 2;
  const row = Math.floor(index / 2);
  const x = 1.2 + (col * 4);
  const y = 1.8 + (row * 1.4);

  slide20.addShape(pptx.ShapeType.roundRect, {
    x: x,
    y: y,
    w: 3.6,
    h: 1.1,
    fill: { color: colors.white, transparency: 15 },
    line: { color: colors.white, width: 1.5 },
    rectRadius: 0.15
  });

  slide20.addText(item.icon, {
    x: x + 0.3,
    y: y + 0.2,
    w: 0.6,
    h: 0.6,
    fontSize: 36
  });

  slide20.addText(item.quarter, {
    x: x + 1.1,
    y: y + 0.2,
    w: 2.2,
    h: 0.25,
    fontSize: 14,
    bold: true,
    color: colors.white
  });

  slide20.addText(item.feature, {
    x: x + 1.1,
    y: y + 0.52,
    w: 2.2,
    h: 0.4,
    fontSize: 13,
    color: colors.white,
    transparency: 20
  });
});

// ==========================================
// SLIDE 21: Contact Information
// ==========================================
const slide21 = pptx.addSlide();

slide21.background = { fill: colors.lightGray };

slide21.addText('Get in Touch', {
  x: 0.5,
  y: 0.8,
  w: 9,
  h: 0.8,
  fontSize: 40,
  bold: true,
  color: colors.primary,
  align: 'center'
});

const contactInfo = [
  { icon: '📧', label: 'Email', value: 'info@mailbooking.com' },
  { icon: '📞', label: 'Phone', value: '+1 (800) 123-4567' },
  { icon: '🌐', label: 'Website', value: 'www.mailbooking.com' },
  { icon: '📍', label: 'Address', value: '123 Business St, Tech City, TC 12345' }
];

contactInfo.forEach((info, index) => {
  const y = 2 + (index * 0.6);

  slide21.addText(info.icon, {
    x: 2,
    y: y,
    w: 0.5,
    h: 0.5,
    fontSize: 28
  });

  slide21.addText(info.label + ':', {
    x: 2.7,
    y: y + 0.05,
    w: 1.5,
    h: 0.4,
    fontSize: 14,
    bold: true,
    color: colors.primary,
    valign: 'middle'
  });

  slide21.addText(info.value, {
    x: 4.2,
    y: y + 0.05,
    w: 4,
    h: 0.4,
    fontSize: 13,
    color: colors.text,
    valign: 'middle'
  });
});

// ==========================================
// SLIDE 22: Thank You
// ==========================================
const slide22 = pptx.addSlide();

slide22.background = {
  fill: colors.primary
};

// Main text
slide22.addText('Thank You', {
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
slide22.addText('Questions & Feedback Welcome', {
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
    console.log('📊 Slides: 22');
    console.log('');
    console.log('Slide Overview:');
    console.log('  1. Title Slide - Main introduction');
    console.log('  2. Mail Booking Features - 6 feature cards');
    console.log('  3. Key Benefits - 4 numbered benefits');
    console.log('  4. Booking Process - 5-step flow');
    console.log('  5. Platform Statistics - Metrics');
    console.log('  6. User Interface Overview');
    console.log('  7. Mobile App Features');
    console.log('  8. Web Portal Features');
    console.log('  9. Admin Dashboard');
    console.log('  10. Payment Gateway Integration');
    console.log('  11. Notification System');
    console.log('  12. Tracking & Monitoring');
    console.log('  13. Security Features');
    console.log('  14. API Integration');
    console.log('  15. Customer Support');
    console.log('  16. Pricing Plans');
    console.log('  17. Implementation Timeline');
    console.log('  18. Success Stories');
    console.log('  19. Technical Architecture');
    console.log('  20. Future Roadmap');
    console.log('  21. Contact Information');
    console.log('  22. Thank You');
  })
  .catch(err => {
    console.error('❌ Error creating presentation:', err);
  });

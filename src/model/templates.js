import { newDeck, newSlide, fillPlaceholders, newShape, textToHtml } from './deck.js';

/** Starter decks on the start screen. `build()` returns a new deck. */
const make = (title, themeId, slides) => () => {
  const deck = newDeck({ title, themeId, slides: [] });
  deck.slides = slides.map(([layout, content, extra]) => {
    const s = fillPlaceholders(newSlide(layout, deck), content || {});
    if (extra) extra(s, deck);
    return s;
  });
  return deck;
};

/** Three coloured stat cards across the lower part of a slide. */
const statCards = (stats) => (slide) => {
  slide.elements = slide.elements.filter((e) => e.ph !== 'body');
  stats.forEach(([value, label], i) => {
    const card = newShape('roundRect', 58 + i * 288, 160, 268, 300);
    card.fill = `theme:accent${i + 1}`;
    card.text = `<p><span style="font-size: 54px; font-weight: 700">${value}</span></p><p>${label}</p>`;
    card.style = { size: 18, color: '#ffffff', align: 'center', valign: 'middle' };
    card.anim = { effect: 'fly-up', order: i + 1, duration: 0.5 };
    slide.elements.push(card);
  });
};

const steps = (labels) => (slide) => {
  labels.forEach((label, i) => {
    const s = newShape('chevron', 58 + i * 214, 230, 220, 110);
    s.fill = `theme:accent${(i % 6) + 1}`;
    s.text = textToHtml(label);
    s.style = { size: 16, color: '#ffffff', align: 'center', valign: 'middle', bold: true };
    s.anim = { effect: 'fade', order: i + 1, duration: 0.5 };
    slide.elements.push(s);
  });
};

export const TEMPLATES = [
  {
    id: 'blank', name: 'Blank presentation',
    build: () => newDeck(),
  },
  {
    id: 'pitch', name: 'Pitch deck',
    build: make('Pitch deck', 'midnight', [
      ['title', { title: 'Company name', subtitle: 'One line that says what you do' }],
      ['titleContent', { title: 'The problem', body: ['Who has the problem, and how often', 'What it costs them today', 'Why existing solutions fall short'] }],
      ['titleContent', { title: 'Our solution', body: ['What the product does', 'The moment customers get value', 'Why now'] }],
      ['bigNumber', { title: '$4.2B', body: 'Serviceable market, growing 18% a year' }],
      ['titleOnly', { title: 'Traction' }, statCards([['12k', 'Weekly active users'], ['38%', 'Month-over-month growth'], ['92', 'Net promoter score']])],
      ['twoContent', { title: 'Business model', body: ['Pricing', 'Sales channel', 'Unit economics'], body2: ['Competition', 'Our advantage', 'Defensibility'] }],
      ['section', { title: 'The ask', subtitle: 'Raising $2M to reach 100k users' }],
    ]),
  },
  {
    id: 'update', name: 'Project update',
    build: make('Project update', 'forest', [
      ['title', { title: 'Project update', subtitle: 'Team name · Month Year' }],
      ['titleContent', { title: 'Highlights', body: ['Shipped the first milestone on schedule', 'Onboarded two new partner teams', 'Cut build times by 40%'] }],
      ['titleOnly', { title: 'Roadmap' }, steps(['Discover', 'Design', 'Build', 'Launch'])],
      ['twoContent', { title: 'Risks and asks', body: ['Hiring is behind plan', 'Vendor contract renewal'], body2: ['Approve two contractor roles', 'Decide on vendor by the 15th'] }],
      ['section', { title: 'Questions?', subtitle: 'team@example.com' }],
    ]),
  },
  {
    id: 'lesson', name: 'Lesson',
    build: make('Lesson', 'editorial', [
      ['title', { title: 'The Water Cycle', subtitle: 'Earth Science · Unit 3' }],
      ['titleContent', { title: 'Today we will', body: ['Describe evaporation, condensation and precipitation', 'Explain how energy from the sun drives the cycle', 'Model the cycle with a simple experiment'] }],
      ['quote', { title: '“Water is the driving force of all nature.”', subtitle: '— Leonardo da Vinci' }],
      ['twoContent', { title: 'Key vocabulary', body: ['Evaporation', 'Condensation', 'Precipitation'], body2: ['Collection', 'Transpiration', 'Runoff'] }],
      ['titleContent', { title: 'Exit ticket', body: ['Draw the water cycle from memory', 'Label each stage', 'Circle the step that needs the most energy'] }],
    ]),
  },
  {
    id: 'workshop', name: 'Workshop',
    build: make('Workshop', 'coral', [
      ['title', { title: 'Design Sprint Workshop', subtitle: 'A five-day process for answering big questions' }],
      ['titleContent', { title: 'Agenda', body: ['Map the problem', 'Sketch competing solutions', 'Decide on the best', 'Prototype', 'Test with real users'] }],
      ['titleOnly', { title: 'How the week flows' }, steps(['Map', 'Sketch', 'Decide', 'Prototype'])],
      ['bigNumber', { title: '5 days', body: 'From big question to tested prototype' }],
      ['section', { title: 'Let’s get started', subtitle: 'Grab a marker and some sticky notes' }],
    ]),
  },
  {
    id: 'announcement', name: 'Announcement',
    build: make('Announcement', 'bold', [
      ['title', { title: 'BIG NEWS', subtitle: 'Something new is coming' }],
      ['bigNumber', { title: '03.14', body: 'Save the date' }],
      ['titleContent', { title: 'What to expect', body: ['A brand-new experience', 'Exclusive early access', 'Surprises along the way'] }],
      ['section', { title: 'SEE YOU THERE', subtitle: 'example.com/launch' }],
    ]),
  },
  {
    id: 'portfolio', name: 'Portfolio',
    build: make('Portfolio', 'plum', [
      ['title', { title: 'Your Name', subtitle: 'Designer · Illustrator · Storyteller' }],
      ['titleContent', { title: 'About me', body: ['Ten years making things people love', 'Based in Lisbon, working everywhere', 'Clients: studios, startups, museums'] }],
      ['titleOnly', { title: 'Selected work' }, statCards([['01', 'Brand system'], ['02', 'Editorial'], ['03', 'Product']])],
      ['section', { title: 'Let’s work together', subtitle: 'hello@example.com' }],
    ]),
  },
];

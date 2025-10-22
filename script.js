const collections = [
  {
    title: 'Powertrain Dismantling Manuals',
    count: 18,
    description: 'Step-by-step guidance for drivetrain removal, inspection, and component harvesting across major OEM platforms.',
    tags: ['Engines', 'Transmissions', 'Axles']
  },
  {
    title: 'Safety & Compliance Briefings',
    count: 12,
    description: 'Interactive toolbox talks and compliance policies covering PPE, hazmat handling, and OSHA reporting.',
    tags: ['Safety', 'Compliance', 'Training']
  },
  {
    title: 'Training & Certifications',
    count: 9,
    description: 'Role-based learning paths, onboarding checklists, and certification tracking for dismantling technicians.',
    tags: ['Onboarding', 'Skill Development', 'HR']
  },
  {
    title: 'Data Dashboards & Reports',
    count: 6,
    description: 'Downloadable dashboards for throughput metrics, part grading performance, and quality KPIs.',
    tags: ['Analytics', 'Supervisors']
  }
];

const featuredItems = [
  {
    title: 'Ford F-150 EcoBoost Teardown Playbook',
    type: 'Manual',
    description:
      'High-resolution walkthrough with torque specs, tooling callouts, and QA checkpoints for the 3.5L EcoBoost platform.',
    meta: ['Updated: March 2024', 'Format: Interactive PDF', 'Author: Powertrain Engineering'],
    link: 'https://example.com/ford-ecoboost-teardown'
  },
  {
    title: 'Lithium-Ion Battery Isolation Drill',
    type: 'Training',
    description:
      'Immersive learning module practicing high-voltage isolation and storage using VR-assisted simulations.',
    meta: ['Duration: 25 minutes', 'Prerequisite: High Voltage Safety'],
    link: 'https://example.com/battery-isolation'
  },
  {
    title: 'Daily Throughput Dashboard',
    type: 'Dashboard',
    description:
      'Live analytics snapshot of dismantle queue status, cycle time, and scrap ratios with exportable CSV reports.',
    meta: ['Data refresh: 15 minutes', 'Owner: Operations Analytics'],
    link: 'https://example.com/throughput-dashboard'
  }
];

const searchIndex = featuredItems.map(item => ({
  ...item,
  searchString: [item.title, item.type, item.description, item.meta.join(' ')].join(' ').toLowerCase()
}));

function renderCollections() {
  const container = document.querySelector('.collection-grid');
  const template = document.querySelector('#collection-card-template');

  collections.forEach(collection => {
    const card = template.content.cloneNode(true);
    card.querySelector('.card-title').textContent = collection.title;
    card.querySelector('.card-count').textContent = `${collection.count} resources`;
    card.querySelector('.card-description').textContent = collection.description;

    const tagsContainer = card.querySelector('.card-tags');
    collection.tags.forEach(tag => {
      const tagElement = document.createElement('span');
      tagElement.className = 'card-tag';
      tagElement.textContent = tag;
      tagsContainer.appendChild(tagElement);
    });

    container.appendChild(card);
  });
}

function renderFeaturedItems(items = featuredItems) {
  const container = document.querySelector('.featured-grid');
  const template = document.querySelector('#featured-card-template');
  container.innerHTML = '';

  items.forEach(item => {
    const card = template.content.cloneNode(true);
    card.querySelector('.featured-title').textContent = item.title;
    card.querySelector('.featured-type').textContent = item.type;
    card.querySelector('.featured-description').textContent = item.description;

    const metaList = card.querySelector('.featured-meta');
    item.meta.forEach(metaItem => {
      const li = document.createElement('li');
      li.textContent = metaItem;
      metaList.appendChild(li);
    });

    const link = card.querySelector('.featured-link');
    link.textContent = 'Open Resource';
    link.href = item.link;

    container.appendChild(card);
  });

  if (items.length === 0) {
    const emptyState = document.createElement('p');
    emptyState.className = 'empty-state';
    emptyState.textContent = 'No matching resources. Try a different keyword or explore the collections above.';
    container.appendChild(emptyState);
  }
}

function handleSearch(event) {
  if (event.type === 'keydown' && event.key !== 'Enter') {
    return;
  }

  const query = document.querySelector('#library-search').value.trim().toLowerCase();
  if (!query) {
    renderFeaturedItems();
    return;
  }

  const filtered = searchIndex.filter(item => item.searchString.includes(query));
  renderFeaturedItems(filtered);
}

function initializeSearch() {
  const searchInput = document.querySelector('#library-search');
  const searchButton = document.querySelector('#search-button');

  searchButton.addEventListener('click', handleSearch);
  searchInput.addEventListener('keydown', handleSearch);
}

function setYear() {
  const yearElement = document.querySelector('#year');
  yearElement.textContent = new Date().getFullYear();
}

document.addEventListener('DOMContentLoaded', () => {
  renderCollections();
  renderFeaturedItems();
  initializeSearch();
  setYear();
});

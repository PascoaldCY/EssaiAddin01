// tableau.js
class TableauExercise {
  constructor() {
      this.currentStep = 1; // Start directly at step 1 since we come from welcome page
      this.isInitialized = true; // Already initialized from welcome page
      
      this.initializeElements();
      this.attachEventListeners();
      this.initializeFromWelcome();
      this.updateProgress();
  }

  initializeElements() {
      // Buttons
      this.validateBtn = document.getElementById('tableau-validate-btn');
      this.closeResultsBtn = document.getElementById('tableau-close-results-btn');

      // Modals
      this.resultsModal = document.getElementById('tableau-results-modal');

      // Steps
      this.stepInit = document.getElementById('tableau-step-init');
      this.stepCustomize = document.getElementById('tableau-step-customize');
      this.stepValidate = document.getElementById('tableau-step-validate');

      // Other elements
      this.status = document.getElementById('tableau-status');
      this.progress = document.getElementById('tableau-progress');
      this.tablePreview = document.getElementById('tableau-table-preview');
      this.validationResult = document.getElementById('tableau-validation-result');
  }

  attachEventListeners() {
      this.validateBtn.addEventListener('click', () => this.handleValidation());
      this.closeResultsBtn.addEventListener('click', () => this.hideModal(this.resultsModal));

      // Close modal on background click
      [this.resultsModal].forEach(modal => {
          modal.addEventListener('click', (e) => {
              if (e.target === modal) {
                  this.hideModal(modal);
              }
          });
      });
  }

  initializeFromWelcome() {
      // Update UI to show exercise is ready
      this.status.textContent = "Exercice initialisé - Prêt pour la personnalisation";
      this.status.style.background = "rgba(40, 167, 69, 0.2)";
      
      // Show the appropriate steps
      this.stepCustomize.classList.remove('hidden');
      this.stepValidate.classList.remove('hidden');
      
      // Generate sample table immediately
      this.generateSampleTable();
      
      // Add animation
      this.animateStep(this.stepCustomize);
  }

  handleInitialization() {
      if (this.hasContent) {
          this.showModal(this.confirmationModal);
      } else {
          this.initializeExercise();
      }
  }

  confirmInitialization() {
      this.hideModal(this.confirmationModal);
      this.initializeExercise();
  }

  initializeExercise() {
      this.isInitialized = true;
      this.currentStep = 1;
      
      // Update UI
      this.status.textContent = "Document initialisé - Prêt pour la personnalisation";
      this.status.style.background = "rgba(40, 167, 69, 0.2)";
      
      // Hide init step, show customize step
      this.stepInit.classList.add('hidden');
      this.stepCustomize.classList.remove('hidden');
      this.stepValidate.classList.remove('hidden');
      
      // Generate sample table
      this.generateSampleTable();
      
      // Update progress
      this.updateProgress();
      
      // Add some animation
      this.animateStep(this.stepCustomize);
  }

  generateSampleTable() {
      const sampleData = [
          ['Produit', 'Catégorie', 'Prix', 'Stock', 'Statut'],
          ['Laptop Dell XPS', 'Informatique', '1299€', '15', 'Disponible'],
          ['iPhone 15 Pro', 'Téléphonie', '1199€', '8', 'Disponible'],
          ['Samsung Galaxy Tab', 'Tablettes', '599€', '12', 'Disponible'],
          ['MacBook Air M2', 'Informatique', '1499€', '5', 'Stock faible'],
          ['AirPods Pro', 'Audio', '279€', '25', 'Disponible'],
          ['Surface Pro 9', 'Informatique', '1099€', '0', 'Rupture']
      ];

      const table = document.createElement('table');
      table.className = 'tableau-sample-table';
      
      sampleData.forEach((row, index) => {
          const tr = document.createElement('tr');
          row.forEach(cell => {
              const cellElement = document.createElement(index === 0 ? 'th' : 'td');
              cellElement.textContent = cell;
              tr.appendChild(cellElement);
          });
          table.appendChild(tr);
      });

      this.tablePreview.innerHTML = '';
      this.tablePreview.appendChild(table);
  }

  handleValidation() {
      // Simulate validation process
      this.validateBtn.disabled = true;
      this.validateBtn.textContent = 'Validation en cours...';
      
      setTimeout(() => {
          this.performValidation();
          this.validateBtn.disabled = false;
          this.validateBtn.textContent = 'Vérifier mon travail';
      }, 2000);
  }

  performValidation() {
      // Simulate different validation scenarios
      const validationResults = this.simulateValidation();
      
      this.currentStep = 2;
      this.updateProgress();
      
      this.displayValidationResults(validationResults);
      this.showModal(this.resultsModal);
  }

  simulateValidation() {
      const scenarios = [
          {
              success: true,
              score: 95,
              message: "Excellent travail ! Votre tableau est parfaitement personnalisé.",
              details: [
                  { type: 'success', text: 'Style personnalisé appliqué correctement' },
                  { type: 'success', text: 'Couleurs et bordures bien choisies' },
                  { type: 'success', text: 'Ligne d\'en-tête correctement formatée' },
                  { type: 'success', text: 'Bandes alternées appliquées' },
                  { type: 'success', text: 'Tableau lisible et esthétique' }
              ]
          },
          {
              success: true,
              score: 80,
              message: "Très bon travail avec quelques améliorations possibles.",
              details: [
                  { type: 'success', text: 'Style personnalisé appliqué correctement' },
                  { type: 'success', text: 'Couleurs bien choisies' },
                  { type: 'warning', text: 'Bordures pourraient être plus visibles' },
                  { type: 'success', text: 'Ligne d\'en-tête correctement formatée' },
                  { type: 'warning', text: 'Alignement du texte à améliorer dans certaines cellules' }
              ]
          },
          {
              success: false,
              score: 60,
              message: "Bon début, mais plusieurs éléments nécessitent des corrections.",
              details: [
                  { type: 'success', text: 'Tableau présent et structuré' },
                  { type: 'error', text: 'Style personnalisé non appliqué ou incomplet' },
                  { type: 'warning', text: 'Couleurs peu contrastées, lisibilité difficile' },
                  { type: 'error', text: 'Ligne d\'en-tête non formatée' },
                  { type: 'warning', text: 'Bandes alternées manquantes' }
              ]
          }
      ];

      return scenarios[Math.floor(Math.random() * scenarios.length)];
  }

  displayValidationResults(results) {
      const title = document.getElementById('tableau-results-title');
      const content = document.getElementById('tableau-results-content');
      
      title.textContent = results.success ? '✅ Validation réussie !' : '⚠️ Validation incomplète';
      
      const scoreColor = results.score >= 80 ? '#27ae60' : results.score >= 60 ? '#f39c12' : '#e74c3c';
      
      content.innerHTML = `
          <div class="tableau-score-display" style="margin-bottom: 20px;">
              <div style="font-size: 2rem; font-weight: bold; color: ${scoreColor};">
                  ${results.score}/100
              </div>
              <p style="margin: 10px 0; font-size: 1.1rem; color: #555;">
                  ${results.message}
              </p>
          </div>
          
          <div class="tableau-validation-details">
              <h4 style="margin-bottom: 15px; color: #2c3e50;">Détails de la validation :</h4>
              <ul class="tableau-validation-list">
                  ${results.details.map(detail => `
                      <li>
                          <span class="tableau-validation-icon ${detail.type}">
                              ${detail.type === 'success' ? '✅' : detail.type === 'warning' ? '⚠️' : '❌'}
                          </span>
                          ${detail.text}
                      </li>
                  `).join('')}
              </ul>
          </div>
          
          ${results.score < 80 ? `
              <div style="margin-top: 20px; padding: 15px; background: #e8f4f8; border-radius: 8px; border-left: 4px solid #3498db;">
                  <h5 style="color: #2980b9; margin-bottom: 10px;">💡 Conseils pour améliorer :</h5>
                  <ul style="text-align: left; margin-left: 20px; color: #2c5282;">
                      <li>Explorez les styles prédéfinis dans l'onglet "Création de tableau"</li>
                      <li>Utilisez des couleurs contrastées pour améliorer la lisibilité</li>
                      <li>Assurez-vous que la ligne d'en-tête se distingue clairement</li>
                      <li>Testez différents alignements pour optimiser la présentation</li>
                  </ul>
              </div>
          ` : ''}
      `;

      // Update main status
      if (results.success && results.score >= 80) {
          this.status.textContent = "Exercice terminé avec succès !";
          this.status.style.background = "rgba(40, 167, 69, 0.3)";
      } else {
          this.status.textContent = "Exercice à améliorer";
          this.status.style.background = "rgba(255, 193, 7, 0.3)";
      }
  }

  updateProgress() {
      const progressPercentage = (this.currentStep / 2) * 100;
      this.progress.style.width = `${progressPercentage}%`;
  }

  showModal(modal) {
      modal.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
  }

  hideModal(modal) {
      modal.classList.add('hidden');
      document.body.style.overflow = 'auto';
  }

  animateStep(step) {
      step.style.opacity = '0';
      step.style.transform = 'translateY(20px)';
      
      setTimeout(() => {
          step.style.transition = 'all 0.5s ease';
          step.style.opacity = '1';
          step.style.transform = 'translateY(0)';
      }, 100);
  }
}

document.addEventListener('DOMContentLoaded', () => {
// MENU ET CONFIRMATION
const startBtn = document.getElementById('tableau-start-btn');
const confirmationModal = document.getElementById('tableau-confirmation-modal');
const cancelBtn = document.getElementById('tableau-cancel-btn');
const confirmBtn = document.getElementById('tableau-confirm-btn');

startBtn.addEventListener('click', () => {
    const tablePreview = document.getElementById('tableau-table-preview');
    const hasContent = tablePreview.innerHTML.trim() !== '';

    if (hasContent) {
        confirmationModal.classList.remove('hidden');
        document.body.style.overflow = 'hidden';
    } else {
        startExercise();
    }
});

cancelBtn.addEventListener('click', () => {
    confirmationModal.classList.add('hidden');
    document.body.style.overflow = 'auto';
});

confirmBtn.addEventListener('click', () => {
    document.getElementById('tableau-table-preview').innerHTML = '';
    confirmationModal.classList.add('hidden');
    document.body.style.overflow = 'auto';
    startExercise();
});

function startExercise() {
    document.getElementById('tableau-menu-step').classList.add('hidden');
    new TableauExercise();  // L'exercice ne démarre que maintenant
}
});


// Add some interactive features
document.addEventListener('DOMContentLoaded', () => {
  // Add hover effects to task list items
  const taskItems = document.querySelectorAll('.tableau-task-list li');
  taskItems.forEach(item => {
      item.addEventListener('mouseenter', () => {
          item.style.backgroundColor = '#e3f2fd';
          item.style.borderLeft = '4px solid #2196f3';
      });
      
      item.addEventListener('mouseleave', () => {
          item.style.backgroundColor = 'white';
          item.style.borderLeft = 'none';
      });
  });

  // Add keyboard navigation
  document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
          const visibleModal = document.querySelector('.tableau-modal:not(.hidden)');
          if (visibleModal) {
              visibleModal.classList.add('hidden');
              document.body.style.overflow = 'auto';
          }
      }
  });

  // Add smooth scrolling for better UX
  const smoothScroll = (target) => {
      target.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
      });
  };

  // Add click handlers for smooth navigation
  document.querySelectorAll('button').forEach(button => {
      button.addEventListener('click', (e) => {
          // Add ripple effect
          const ripple = document.createElement('span');
          ripple.classList.add('ripple');
          button.appendChild(ripple);
          
          const rect = button.getBoundingClientRect();
          const size = Math.max(rect.width, rect.height);
          const x = e.clientX - rect.left - size / 2;
          const y = e.clientY - rect.top - size / 2;
          
          ripple.style.width = ripple.style.height = size + 'px';
          ripple.style.left = x + 'px';
          ripple.style.top = y + 'px';
          
          setTimeout(() => {
              if (ripple.parentNode) {
                  ripple.parentNode.removeChild(ripple);
              }
          }, 600);
      });
  });
});

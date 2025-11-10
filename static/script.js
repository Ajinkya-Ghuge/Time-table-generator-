// Plannable.ai Custom JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // File upload preview enhancement
    const fileInputs = document.querySelectorAll('input[type="file"]');
    
    fileInputs.forEach(input => {
        input.addEventListener('change', function(e) {
            const fileName = this.files[0]?.name;
            if (fileName) {
                const label = this.previousElementSibling;
                if (label && label.classList.contains('form-label')) {
                    const originalText = label.textContent;
                    label.innerHTML = `${originalText} <span class="badge bg-success ms-2">${fileName}</span>`;
                }
            }
        });
    });

    // Form submission loading state
    const forms = document.querySelectorAll('form');
    forms.forEach(form => {
        form.addEventListener('submit', function() {
            const submitBtn = this.querySelector('button[type="submit"]');
            if (submitBtn) {
                submitBtn.disabled = true;
                submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Processing...';
            }
        });
    });

    // Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert');
    alerts.forEach(alert => {
        setTimeout(() => {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });

    // Smooth scrolling for anchor links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // Tab functionality enhancement
    const tabPanes = document.querySelectorAll('.tab-pane');
    tabPanes.forEach(pane => {
        pane.addEventListener('show.bs.tab', function() {
            // Add any custom tab show logic here
        });
    });

    // Download button enhancement
    const downloadBtn = document.querySelector('a[href="/download_timetable"]');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', function() {
            this.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Preparing Download...';
        });
    }

    // Add copy to clipboard functionality for timetable cells
    const timetableCells = document.querySelectorAll('.timetable-table td');
    timetableCells.forEach(cell => {
        cell.addEventListener('click', function() {
            const text = this.textContent.trim();
            if (text && !text.includes('Break')) {
                navigator.clipboard.writeText(text).then(() => {
                    // Show temporary tooltip
                    const originalText = this.textContent;
                    this.textContent = 'Copied!';
                    setTimeout(() => {
                        this.textContent = originalText;
                    }, 1000);
                });
            }
        });
    });

    // Responsive table handling
    function handleResponsiveTables() {
        const tables = document.querySelectorAll('.table-responsive');
        tables.forEach(table => {
            if (table.offsetWidth < table.scrollWidth) {
                table.parentElement.classList.add('table-scroll-indicator');
            }
        });
    }

    handleResponsiveTables();
    window.addEventListener('resize', handleResponsiveTables);
});

// Utility functions
const PlannableUtils = {
    // Format file size
    formatFileSize: function(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },

    // Validate CSV file
    validateCSVFile: function(file) {
        const validTypes = ['text/csv', 'application/vnd.ms-excel'];
        const maxSize = 5 * 1024 * 1024; // 5MB
        
        if (!validTypes.includes(file.type) && !file.name.endsWith('.csv')) {
            return 'Please upload a valid CSV file.';
        }
        
        if (file.size > maxSize) {
            return 'File size must be less than 5MB.';
        }
        
        return null;
    },

    // Show notification
    showNotification: function(message, type = 'info') {
        const alertClass = {
            'success': 'alert-success',
            'error': 'alert-danger',
            'warning': 'alert-warning',
            'info': 'alert-info'
        }[type] || 'alert-info';

        const notification = document.createElement('div');
        notification.className = `alert ${alertClass} alert-dismissible fade show`;
        notification.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            if (notification.parentElement) {
                notification.remove();
            }
        }, 5000);
    }
};
document.addEventListener('DOMContentLoaded', () => {
    // --- Application State ---
    let currentView = 'dashboard';

    // --- DOM Elements ---
    const sidebar = document.getElementById('sidebar');
    const mobileToggle = document.getElementById('mobileToggle');
    const mobileToggleHeader = document.getElementById('mobileToggleHeader');
    const navItems = document.querySelectorAll('.nav-item');
    const views = document.querySelectorAll('.view');
    const pageTitle = document.getElementById('pageTitle');

    // Modals
    const modals = {
        service: document.getElementById('modalService'),
        wo: document.getElementById('modalWO'),
        financial: document.getElementById('modalFinancial'),
        confirm: document.getElementById('modalConfirm')
    };

    // Forms
    const forms = {
        service: document.getElementById('formService'),
        wo: document.getElementById('formWO'),
        financial: document.getElementById('formFinancial')
    };

    // Global Chart Instance
    let successChartInstance = null;
    
    // Deletion State
    let pendingDeletion = { type: null, id: null };
    
    // Formatter
    const currencyFormatter = new Intl.NumberFormat('pt-PT', { style: 'currency', currency: 'EUR' });
    const dateFormatter = new Intl.DateTimeFormat('pt-BR');

    // --- Initialization ---

    function init() {
        setupNavigation();
        setupModals();
        setupForms();
        setupFilters();
        
        // Load initial view
        switchView('dashboard');
        
        // Wait for Dexie to be ready if it's seeding data
        db.on('ready', () => {
            refreshCurrentView();
        });
        // Also call now in case it's already ready
        refreshCurrentView();
    }

    // --- Navigation ---

    function setupNavigation() {
        const toggleMenu = () => {
            const isActive = sidebar.classList.toggle('active');
            document.body.classList.toggle('sidebar-open');
            
            // Switch icons if header toggle exists
            if (mobileToggleHeader) {
                const icon = mobileToggleHeader.querySelector('i');
                if (icon) {
                    icon.className = isActive ? 'fa-solid fa-xmark' : 'fa-solid fa-bars';
                }
            }
        };

        mobileToggle.addEventListener('click', toggleMenu);
        if (mobileToggleHeader) {
            mobileToggleHeader.addEventListener('click', toggleMenu);
        }

        navItems.forEach(item => {
            item.addEventListener('click', (e) => {
                e.preventDefault();
                const viewTarget = e.currentTarget.getAttribute('data-view');
                switchView(viewTarget);
                
                // Close sidebar on mobile after click
                if (window.innerWidth <= 768) {
                    sidebar.classList.remove('active');
                    document.body.classList.remove('sidebar-open');
                }
            });
        });
    }

    function switchView(viewId) {
        currentView = viewId;
        
        // Update Navigation
        navItems.forEach(item => {
            item.classList.remove('active');
            if(item.getAttribute('data-view') === viewId) {
                item.classList.add('active');
                pageTitle.textContent = item.querySelector('span').textContent;
            }
        });

        // Update Views
        views.forEach(view => {
            view.classList.remove('active');
            if(view.id === `view-${viewId}`) {
                view.classList.add('active');
            }
        });

        // Load data for view
        refreshCurrentView();
    }

    function refreshCurrentView() {
        if(currentView === 'dashboard') loadDashboard();
        if(currentView === 'services') loadServices();
        if(currentView === 'work-orders') loadWorkOrders();
        if(currentView === 'financial') loadFinancials();
    }

    // --- Modals & Global UI ---

    function setupModals() {
        // Setup Close buttons
        document.querySelectorAll('.close-modal, .cancel-modal').forEach(btn => {
            btn.addEventListener('click', closeAllModals);
        });

        // Define specific modal triggers
        document.getElementById('btnNewService').addEventListener('click', () => openModal('service'));
        document.getElementById('btnNewWO').addEventListener('click', () => openModal('wo'));
        document.getElementById('btnNewIncome').addEventListener('click', () => openModal('financial', { type: 'entrada' }));
        document.getElementById('btnNewExpense').addEventListener('click', () => openModal('financial', { type: 'saida' }));
        
        // Confirm Modal Action
        document.getElementById('btnConfirmAction').addEventListener('click', executeDeletion);
    }

    async function openModal(modalName, data = null) {
        closeAllModals();
        const modal = modals[modalName];
        if(!modal) return;
        
        if (modalName === 'service') {
            document.getElementById('modalServiceTitle').textContent = data ? 'Editar Serviço' : 'Novo Serviço';
            document.getElementById('serviceId').value = data ? data.id : '';
            document.getElementById('serviceName').value = data ? data.name : '';
            document.getElementById('serviceDesc').value = data ? data.description : '';
        } 
        else if (modalName === 'wo') {
            document.getElementById('modalWOTitle').textContent = data ? 'Editar Ordem de Serviço' : 'Nova Ordem de Serviço';
            document.getElementById('woId').value = data ? data.id : '';
            document.getElementById('woCode').value = data ? data.wo : '';
            document.getElementById('woCode').disabled = data ? true : false; // Can't edit WO ID
            
            // Populate Services dropdown
            const services = await db.services.toArray();
            const select = document.getElementById('woService');
            select.innerHTML = '<option value="">Selecione...</option>';
            services.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s.id;
                opt.textContent = s.name;
                select.appendChild(opt);
            });

            document.getElementById('woService').value = data ? data.service_id : '';
            document.getElementById('woDate').value = data ? data.date : new Date().toISOString().split('T')[0];
            document.getElementById('woStatus').value = data ? data.status : 'Sucesso';
            document.getElementById('woValue').value = data ? (data.value || '') : '';
            document.getElementById('woErrorMsg').style.display = 'none';
        }
        else if (modalName === 'financial') {
            // Setup Type (Entrada vs Saída)
            const type = data.isEdit ? data.type : (data.type || 'entrada');
            document.getElementById('finType').value = type;
            document.getElementById('modalFinancialTitle').textContent = data.isEdit 
                ? 'Editar Registro' 
                : (type === 'entrada' ? 'Nova Entrada' : 'Nova Saída');
            
            document.getElementById('finId').value = data.isEdit ? data.id : '';
            document.getElementById('finDate').value = data.isEdit ? data.date : new Date().toISOString().split('T')[0];
            document.getElementById('finDesc').value = data.isEdit ? data.description : '';
            document.getElementById('finValue').value = data.isEdit ? data.value : '';
            
            // Populate financial service dropdown
            const services = await db.services.toArray();
            const select = document.getElementById('finService');
            select.innerHTML = '<option value="">Nenhum</option>';
            services.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s.id;
                opt.textContent = s.name;
                select.appendChild(opt);
            });
            
            document.getElementById('finService').value = data.isEdit ? (data.service_id || '') : '';
            
            // Highlight color based on type
            const header = modal.querySelector('.modal-header');
            header.style.backgroundColor = type === 'entrada' ? 'var(--success-color)' : 'var(--danger-color)';
            header.style.color = 'white';
            modal.querySelector('.close-modal').style.color = 'white';
        }

        modal.classList.add('active');
    }

    function closeAllModals() {
        Object.values(modals).forEach(m => m.classList.remove('active'));
    }

    function confirmDelete(type, id, message) {
        pendingDeletion = { type, id };
        document.getElementById('confirmMessage').textContent = message || 'Tem certeza que deseja excluir este item?';
        modals.confirm.classList.add('active');
    }

    async function executeDeletion() {
        const { type, id } = pendingDeletion;
        if (!type || !id) return;
        
        try {
            if (type === 'service') await db.services.delete(parseInt(id));
            if (type === 'wo') await db.workOrders.delete(parseInt(id));
            if (type === 'financial') await db.financials.delete(parseInt(id));
            
            closeAllModals();
            refreshCurrentView();
        } catch (error) {
            alert("Erro ao excluir: " + error.message);
        }
    }

    // --- Forms specific setup ---
    
    function setupForms() {
        forms.service.addEventListener('submit', async (e) => {
            e.preventDefault();
            const id = document.getElementById('serviceId').value;
            const data = {
                name: document.getElementById('serviceName').value,
                description: document.getElementById('serviceDesc').value
            };

            if (id) {
                await db.services.update(parseInt(id), data);
            } else {
                await db.services.add(data);
            }

            closeAllModals();
            loadServices();
        });

        forms.wo.addEventListener('submit', async (e) => {
            e.preventDefault();
            const id = document.getElementById('woId').value;
            const woCode = document.getElementById('woCode').value;
            const woStatus = document.getElementById('woStatus').value;
            const woValueInput = document.getElementById('woValue').value;
            const woValue = woValueInput ? parseFloat(woValueInput) : 0;
            const serviceId = parseInt(document.getElementById('woService').value);
            const dateStr = document.getElementById('woDate').value;
            
            const data = {
                wo: woCode,
                service_id: serviceId,
                date: dateStr,
                status: woStatus,
                value: woValue
            };

            try {
                if (id) {
                    await db.workOrders.update(parseInt(id), data);
                } else {
                    // Check logic manually just to be safe before add
                    const exists = await db.workOrders.where('wo').equals(woCode).count();
                    if(exists > 0) {
                        const err = document.getElementById('woErrorMsg');
                        err.textContent = "Work Order (WO) já existe no sistema!";
                        err.style.display = 'block';
                        return; // Prevent submit
                    }
                    const newWoId = await db.workOrders.add(data);
                    
                    // Add to Financial if greater than 0
                    if (woValue > 0) {
                        const serviceName = document.getElementById('woService').options[document.getElementById('woService').selectedIndex].text;
                        await db.financials.add({
                            type: 'entrada',
                            date: dateStr,
                            description: `Pagamento SO #${woCode} - ${serviceName}`,
                            value: woValue,
                            service_id: serviceId
                        });
                    }
                }
                
                closeAllModals();
                loadWorkOrders();
            } catch (err) {
                document.getElementById('woErrorMsg').textContent = err.message;
                document.getElementById('woErrorMsg').style.display = 'block';
            }
        });

        forms.financial.addEventListener('submit', async (e) => {
            e.preventDefault();
            const id = document.getElementById('finId').value;
            const serviceVal = document.getElementById('finService').value;
            
            const data = {
                type: document.getElementById('finType').value,
                date: document.getElementById('finDate').value,
                description: document.getElementById('finDesc').value,
                value: parseFloat(document.getElementById('finValue').value),
                service_id: serviceVal ? parseInt(serviceVal) : null
            };

            if (id) {
                await db.financials.update(parseInt(id), data);
            } else {
                await db.financials.add(data);
            }

            closeAllModals();
            loadFinancials();
        });
    }

    // --- Helper Functions ---
    function isDateInPeriod(dateStr, period, customStart = null, customEnd = null) {
        if (period === 'all') return true;
        if (!dateStr) return false;
        
        const date = new Date(dateStr);
        const today = new Date();
        today.setHours(0,0,0,0);
        
        const startOfWeek = new Date(today);
        startOfWeek.setDate(today.getDate() - today.getDay());
        
        switch (period) {
            case 'today':
                return date >= today;
            case 'week':
                return date >= startOfWeek;
            case 'month':
                return date.getMonth() === today.getMonth() && date.getFullYear() === today.getFullYear();
            case 'year':
                return date.getFullYear() === today.getFullYear();
            case 'custom':
                if (customStart && customEnd) {
                    return date >= new Date(customStart) && date <= new Date(customEnd);
                } else if (customStart) {
                    return date >= new Date(customStart);
                } else if (customEnd) {
                    return date <= new Date(customEnd);
                }
                return true; // if no dates provided, show all
            default:
                return true;
        }
    }

    // --- View Specific Data Loaders ---

    async function loadDashboard() {
        const period = document.getElementById('dashPeriodFilter').value;
        const customGroup = document.getElementById('dashCustomDateGroup');
        
        // Toggle custom date inputs visibility
        if (period === 'custom') {
            customGroup.style.display = 'flex';
        } else {
            customGroup.style.display = 'none';
        }

        const customStart = document.getElementById('dashStartDate').value;
        const customEnd = document.getElementById('dashEndDate').value;
        
        // Fetch all WOs
        let wos = await db.workOrders.toArray();
        let fins = await db.financials.toArray();
        
        // Apply Period Filter
        wos = wos.filter(wo => isDateInPeriod(wo.date, period, customStart, customEnd));
        fins = fins.filter(fin => isDateInPeriod(fin.date, period, customStart, customEnd));
        
        // --- Calculate KPIs ---
        const totalWOs = wos.length;
        const successWOs = wos.filter(w => w.status === 'Sucesso').length;
        const failWOs = wos.filter(w => w.status === 'Sem Sucesso').length;
        const successRate = totalWOs > 0 ? ((successWOs / totalWOs) * 100).toFixed(1) : 0;

        let totalEntradas = 0;
        let totalSaidas = 0;
        fins.forEach(f => {
            if (f.type === 'entrada') totalEntradas += parseFloat(f.value);
            if (f.type === 'saida') totalSaidas += parseFloat(f.value);
        });

        const saldo = totalEntradas - totalSaidas;

        // Render values
        document.getElementById('kpiEntradas').textContent = currencyFormatter.format(totalEntradas);
        document.getElementById('kpiSaidas').textContent = currencyFormatter.format(totalSaidas);
        document.getElementById('kpiSaldo').textContent = currencyFormatter.format(saldo);
        
        document.getElementById('kpiSaldo').className = 'kpi-value ' + (saldo >= 0 ? 'text-green' : 'text-red');

        document.getElementById('kpiTotalServicos').textContent = totalWOs;
        document.getElementById('kpiTaxaSucesso').textContent = `${successRate}%`;

        // --- Render Chart ---
        const ctx = document.getElementById('successChart').getContext('2d');
        
        if(successChartInstance) {
            successChartInstance.destroy();
        }

        successChartInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Sucesso', 'Sem Sucesso'],
                datasets: [{
                    data: [successWOs, failWOs],
                    backgroundColor: ['#4CAF50', '#F44336'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom' }
                }
            }
        });

        // --- Render Recent Orders ---
        const listEl = document.getElementById('dashRecentOrders');
        listEl.innerHTML = '';
        
        const recentWOs = await db.workOrders.reverse().limit(5).toArray();
        const services = await db.services.toArray();
        const svcMap = {};
        services.forEach(s => svcMap[s.id] = s.name);

        if(recentWOs.length === 0) {
            listEl.innerHTML = '<li>Nenhuma ordem recente</li>';
        } else {
            recentWOs.forEach(wo => {
                const svcName = svcMap[wo.service_id] || 'Desconhecido';
                const isSuccess = wo.status === 'Sucesso';
                
                const li = document.createElement('li');
                li.className = 'recent-item';
                li.innerHTML = `
                    <div class="recent-item-info">
                        <strong>WO: ${wo.wo} - ${svcName}</strong>
                        <small>${dateFormatter.format(new Date(wo.date))} | Status: ${wo.status}</small>
                    </div>
                    <span class="status-badge ${isSuccess ? 'status-sucesso' : 'status-semsucesso'}">
                        ${isSuccess ? '<i class="fa-solid fa-check"></i>' : '<i class="fa-solid fa-xmark"></i>'}
                    </span>
                `;
                listEl.appendChild(li);
            });
        }
    }

    async function loadServices() {
        const tbody = document.getElementById('servicesTableBody');
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;">Carregando...</td></tr>';
        
        const services = await db.services.toArray();
        tbody.innerHTML = '';
        
        if(services.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;">Nenhum serviço cadastrado.</td></tr>';
            return;
        }

        services.forEach(s => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${s.name}</strong></td>
                <td>${s.description || '-'}</td>
                <td class="actions-col">
                    <button class="btn-icon edit" data-id="${s.id}"><i class="fa-solid fa-pen"></i></button>
                    <button class="btn-icon delete" data-id="${s.id}"><i class="fa-solid fa-trash"></i></button>
                </td>
            `;
            tbody.appendChild(tr);
        });

        // Attach event listeners dynamically
        tbody.querySelectorAll('.edit').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = parseInt(e.currentTarget.getAttribute('data-id'));
                const s = await db.services.get(id);
                openModal('service', s);
            });
        });
        
        tbody.querySelectorAll('.delete').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                confirmDelete('service', id, 'Tem certeza que deseja excluir o serviço? Se ele foi usado em uma OS, haverá inconsistência de dados (A OS mostrará um serviço Desconhecido). Em futuras melhorias o sistema poderá bloquear esta ação.');
            });
        });
    }

    async function loadWorkOrders() {
        // Pre-fill filter dropdown
        const svcs = await db.services.toArray();
        const svcsSelect = document.getElementById('filterWOService');
        
        // Only populate if empty (don't overwrite selection state)
        if(svcsSelect.options.length <= 1) {
            svcs.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s.id;
                opt.textContent = s.name;
                svcsSelect.appendChild(opt);
            });
        }

        // Apply visual filters logic
        const searchTerm = document.getElementById('filterWOSearch').value.toLowerCase();
        const dateFilter = document.getElementById('filterWODate').value;
        const svcFilter = document.getElementById('filterWOService').value; // is string id
        const statusFilter = document.getElementById('filterWOStatus').value;

        let wos = await db.workOrders.reverse().toArray();
        const tbody = document.getElementById('woTableBody');
        tbody.innerHTML = '';

        // Apply filters
        wos = wos.filter(wo => {
            if(searchTerm && !wo.wo.toLowerCase().includes(searchTerm)) return false;
            if(dateFilter && wo.date !== dateFilter) return false;
            if(svcFilter && wo.service_id !== parseInt(svcFilter)) return false;
            if(statusFilter && wo.status !== statusFilter) return false;
            return true;
        });

        // Build Service Map
        const svcMap = {};
        svcs.forEach(s => svcMap[s.id] = s.name);

        if(wos.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">Nenhuma Ordem de Serviço encontrada.</td></tr>';
            return;
        }

        wos.forEach(wo => {
            const svcName = svcMap[wo.service_id] || 'Serviço Excluído/Desconhecido';
            const tr = document.createElement('tr');
            
            tr.innerHTML = `
                <td><strong>${wo.wo}</strong></td>
                <td>${dateFormatter.format(new Date(wo.date))}</td>
                <td>${svcName}</td>
                <td>${wo.value ? currencyFormatter.format(wo.value) : '-'}</td>
                <td>
                    <span class="status-badge ${wo.status === 'Sucesso' ? 'status-sucesso' : 'status-semsucesso'}">
                        ${wo.status}
                    </span>
                </td>
                <td class="actions-col">
                    <button class="btn-icon edit" data-id="${wo.id}"><i class="fa-solid fa-pen"></i></button>
                    <button class="btn-icon delete" data-id="${wo.id}"><i class="fa-solid fa-trash"></i></button>
                </td>
            `;
            tbody.appendChild(tr);
        });

        // Events
        tbody.querySelectorAll('.edit').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = parseInt(e.currentTarget.getAttribute('data-id'));
                const wo = await db.workOrders.get(id);
                openModal('wo', wo);
            });
        });
        
        tbody.querySelectorAll('.delete').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                confirmDelete('wo', id, `Excluir a Ordem de Serviço?`);
            });
        });
    }

    async function loadFinancials() {
        const typeFilter = document.getElementById('filterFinType').value;
        const periodFilter = document.getElementById('filterFinPeriod').value;
        const customGroup = document.getElementById('finCustomDateGroup');

        // Toggle custom date inputs visibility
        if (periodFilter === 'custom') {
            customGroup.style.display = 'flex';
        } else {
            customGroup.style.display = 'none';
        }

        const customStart = document.getElementById('finStartDate').value;
        const customEnd = document.getElementById('finEndDate').value;
        
        let fins = await db.financials.reverse().toArray();
        
        // Apply Filters
        if(typeFilter) {
            fins = fins.filter(f => f.type === typeFilter);
        }
        fins = fins.filter(f => isDateInPeriod(f.date, periodFilter, customStart, customEnd));

        const tbody = document.getElementById('financialTableBody');
        tbody.innerHTML = '';

        if(fins.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">Nenhum registro encontrado.</td></tr>';
            return;
        }

        fins.forEach(f => {
            const isEntrada = f.type === 'entrada';
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${dateFormatter.format(new Date(f.date))}</td>
                <td>
                    <span class="status-badge ${isEntrada ? 'status-sucesso' : 'status-semsucesso'}">
                        <i class="fa-solid fa-${isEntrada ? 'arrow-up' : 'arrow-down'}"></i> ${isEntrada ? 'Entrada' : 'Saída'}
                    </span>
                </td>
                <td><strong>${f.description}</strong></td>
                <td class="${isEntrada ? 'text-green' : 'text-red'} font-weight-bold">
                    ${isEntrada ? '+' : '-'} ${currencyFormatter.format(f.value)}
                </td>
                <td class="actions-col">
                    <button class="btn-icon edit" data-id="${f.id}"><i class="fa-solid fa-pen"></i></button>
                    <button class="btn-icon delete" data-id="${f.id}"><i class="fa-solid fa-trash"></i></button>
                </td>
            `;
            tbody.appendChild(tr);
        });

        // Events
        tbody.querySelectorAll('.edit').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = parseInt(e.currentTarget.getAttribute('data-id'));
                const fin = await db.financials.get(id);
                fin.isEdit = true;
                openModal('financial', fin);
            });
        });
        
        tbody.querySelectorAll('.delete').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-id');
                confirmDelete('financial', id, `Excluir o registro financeiro?`);
            });
        });
    }

    function setupFilters() {
        // WO Filters debounce mapping
        const woInputs = ['filterWOSearch', 'filterWODate', 'filterWOService', 'filterWOStatus'];
        woInputs.forEach(id => {
            const el = document.getElementById(id);
            if(el) {
                el.addEventListener('input', () => {
                    if(currentView === 'work-orders') loadWorkOrders();
                });
            }
        });

        // Dashboard Filters
        document.getElementById('dashPeriodFilter').addEventListener('change', () => {
             loadDashboard();
        });
        document.getElementById('btnDashFilterDate').addEventListener('click', () => {
             loadDashboard();
        });

        // Financial filters
        document.getElementById('filterFinType').addEventListener('change', () => {
             loadFinancials();
        });
        document.getElementById('filterFinPeriod').addEventListener('change', () => {
             loadFinancials();
        });
        document.getElementById('btnFinFilterDate').addEventListener('click', () => {
             loadFinancials();
        });

        // Export functionality via SheetJS
        document.getElementById('btnExportData').addEventListener('click', async (e) => {
            const btn = e.currentTarget;
            const originalContent = btn.innerHTML;
            
            try {
                // UI Feedback: Loading state
                btn.disabled = true;
                btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processando...';
                
                // Fetch data from IndexedDB
                const wos = await db.workOrders.toArray();
                const fins = await db.financials.toArray();
                const svcs = await db.services.toArray();
                const svcMap = {};
                svcs.forEach(s => svcMap[s.id] = s.name);

                // 1. Prepare Work Orders Sheet
                const dataWOs = wos.map(wo => ({
                    'ID OS': wo.id,
                    'Código WO': wo.wo,
                    'Data': wo.date,
                    'Serviço': svcMap[wo.service_id] || 'Desconhecido',
                    'Valor (€)': wo.value || 0,
                    'Status': wo.status
                }));

                // 2. Prepare Financial Sheet
                const dataFins = fins.map(f => ({
                    'ID Reg.': f.id,
                    'Data': f.date,
                    'Tipo': f.type === 'entrada' ? 'Entrada' : 'Saída',
                    'Descrição': f.description,
                    'Valor (€)': f.value,
                    'Serviço Relacionado': f.service_id ? svcMap[f.service_id] : 'Nenhum'
                }));

                // 3. Create Workbook and Worksheets
                const wb = XLSX.utils.book_new();
                const wsWOs = XLSX.utils.json_to_sheet(dataWOs.length > 0 ? dataWOs : [{'Aviso': 'Nenhuma ordem de serviço cadastrada'}]);
                XLSX.utils.book_append_sheet(wb, wsWOs, "Ordens de Serviço");
                const wsFins = XLSX.utils.json_to_sheet(dataFins.length > 0 ? dataFins : [{'Aviso': 'Nenhum registro financeiro cadastrado'}]);
                XLSX.utils.book_append_sheet(wb, wsFins, "Financeiro");

                // 4. Generate file
                const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                const fileName = "CRMTecnico_Export.xlsx";
                const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                const blob = new Blob([wbout], { type: mimeType });

                // 5. Trigger download/share
                // Check for Web Share API support (Mobile priority)
                if (navigator.canShare && navigator.share) {
                    const file = new File([blob], fileName, { type: mimeType });
                    if (navigator.canShare({ files: [file] })) {
                        await navigator.share({
                            files: [file],
                            title: 'Exportação CRM Técnico',
                            text: 'Planilha de Ordens de Serviço e Financeiro'
                        });
                        return; // Successfully shared
                    }
                }

                // Fallback: Desktop Download
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.style.display = 'none';
                a.href = url;
                a.download = fileName;
                document.body.appendChild(a);
                a.click();
                
                setTimeout(() => {
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                }, 100);

            } catch (err) {
                console.error("Export Error:", err);
                if (err.name !== 'AbortError') { // Don't alert if user cancelled share sheet
                    alert("Ocorreu um erro ao exportar os dados: " + err.message);
                }
            } finally {
                // Restore button state
                btn.disabled = false;
                btn.innerHTML = originalContent;
            }
        });
    }

    // Bootstrap
    init();
});

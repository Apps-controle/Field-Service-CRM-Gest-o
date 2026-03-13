// Initialize Dexie local database
const db = new Dexie('CRM Tecnico_DB');

// Define database schema
db.version(2).stores({
    services: '++id, name, description', // Primary key and indexed props
    workOrders: '++id, wo, service_id, date, status, value', // wo is unique but handled via logic
    financials: '++id, type, description, value, date, service_id'
});

// --- Data Access Methods ---

const DBAPI = {
    // Services
    async getServices() {
        return await db.services.toArray();
    },
    async getServiceById(id) {
        return await db.services.get(id);
    },
    async addService(service) {
        return await db.services.add(service);
    },
    async updateService(id, service) {
        return await db.services.update(id, service);
    },
    async deleteService(id) {
        // Option: Check if there are related WOs before deleting
        const relatedWOs = await db.workOrders.where({service_id: id}).count();
        if (relatedWOs > 0) {
            throw new Error('Não é possível excluir: existem Ordens de Serviço vinculadas a este serviço.');
        }
        return await db.services.delete(id);
    },

    // Work Orders
    async getWorkOrders() {
        // Fetch WOs and populate service name
        const wos = await db.workOrders.toArray();
        const services = await this.getServices();
        const serviceMap = {};
        services.forEach(s => serviceMap[s.id] = s.name);
        
        return wos.map(wo => ({
            ...wo,
            serviceName: serviceMap[wo.service_id] || 'Serviço Excluído/Desconhecido'
        }));
    },
    async addWorkOrder(wo) {
        // Check WO uniqueness
        const existing = await db.workOrders.where('wo').equals(wo.wo).first();
        if (existing) {
            throw new Error(`Ordem de serviço ${wo.wo} já existe!`);
        }
        return await db.workOrders.add(wo);
    },
    async updateWorkOrder(id, wo) {
        if(wo.wo) {
            const existing = await db.workOrders.where('wo').equals(wo.wo).first();
            if (existing && existing.id !== id) {
                throw new Error(`Ordem de serviço ${wo.wo} já existe!`);
            }
        }
        return await db.workOrders.update(id, wo);
    },
    async deleteWorkOrder(id) {
        return await db.workOrders.delete(id);
    },

    // Financials
    async getFinancials() {
        return await db.financials.toArray();
    },
    async addFinancial(fin) {
        return await db.financials.add(fin);
    },
    async updateFinancial(id, fin) {
        return await db.financials.update(id, fin);
    },
    async deleteFinancial(id) {
        return await db.financials.delete(id);
    },

    // Dashboard Indicators
    async getDashboardData(period = 'all') { // 'all', 'month', etc (simplified for MVP)
        const wos = await db.workOrders.toArray();
        const fins = await db.financials.toArray();
        
        // Basic filtering can be applied here based on 'period' if needed
        // For now, doing it in memory or passing 'all'
        
        // Service Stats
        const totalWOs = wos.length;
        const successWOs = wos.filter(w => w.status === 'Sucesso').length;
        const failWOs = wos.filter(w => w.status === 'Sem Sucesso').length;
        const successRate = totalWOs > 0 ? ((successWOs / totalWOs) * 100).toFixed(1) : 0;

        // Financial Stats
        let totalEntradas = 0;
        let totalSaidas = 0;
        
        fins.forEach(f => {
            if (f.type === 'entrada') totalEntradas += parseFloat(f.value);
            if (f.type === 'saida') totalSaidas += parseFloat(f.value);
        });

        const saldo = totalEntradas - totalSaidas;

        return {
            services: {
                total: totalWOs,
                success: successWOs,
                fail: failWOs,
                rate: successRate
            },
            financials: {
                entradas: totalEntradas,
                saidas: totalSaidas,
                saldo: saldo
            },
            recentWOs: await this.getRecentWorkOrders(5)
        };
    },

    async getRecentWorkOrders(limit_num) {
        const wos = await db.workOrders.reverse().limit(limit_num).toArray();
        const services = await this.getServices();
        const serviceMap = {};
        services.forEach(s => serviceMap[s.id] = s.name);
        
        return wos.map(wo => ({
            ...wo,
            serviceName: serviceMap[wo.service_id] || 'Desconhecido'
        }));
    }
};

// Seed initial data if empty
db.on('ready', async () => {
    const count = await db.services.count();
    if (count === 0) {
        console.log("Seeding initial data...");
        const initialServices = [
            { name: 'Instalação', description: 'Instalação padrão' },
            { name: 'Suporte', description: 'Atendimento de suporte' },
            { name: 'Adicional', description: 'Serviços adicionais' },
            { name: 'Instalação Amigo', description: 'Promoção amigo' },
            { name: 'Instalação Dados', description: 'Apenas dados' },
            { name: 'Troca de instalação', description: 'Mudança de endereço/equipamento' },
            { name: 'Suporte Puxado Drop', description: 'Manutenção de drop' }
        ];
        await db.services.bulkAdd(initialServices);
    }
});

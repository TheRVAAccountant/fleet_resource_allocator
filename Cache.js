/**
 * ===================================================================
 * CACHE SERVICE
 * ===================================================================
 * High-performance caching layer for expensive operations
 */

class Cache {
  static CACHE_TYPES = {
    SCRIPT: 'script',
    USER: 'user',
    DOCUMENT: 'document'
  };
  
  static stats = {
    hits: 0,
    misses: 0
  };
  
  static get(key, fetcher, options = {}) {
    const {
      ttl = 300,
      type = this.CACHE_TYPES.SCRIPT,
      force = false
    } = options;
    
    if (!force) {
      const cached = this.retrieve(key, type);
      if (cached !== null) {
        this.stats.hits++;
        Logger.createLogger('Cache').debug(`Cache hit for key: ${key}`);
        return cached;
      }
    }
    
    this.stats.misses++;
    Logger.createLogger('Cache').debug(`Cache miss for key: ${key}`);
    const value = fetcher();
    this.store(key, value, ttl, type);
    return value;
  }
  
  static retrieve(key, type) {
    try {
      const cache = this.getCacheService(type);
      const cached = cache.get(key);
      return cached ? JSON.parse(cached) : null;
    } catch (e) {
      Logger.createLogger('Cache').error('Cache retrieve error', { key, error: e.message });
      return null;
    }
  }
  
  static store(key, value, ttl, type) {
    try {
      const cache = this.getCacheService(type);
      cache.put(key, JSON.stringify(value), ttl);
    } catch (e) {
      Logger.createLogger('Cache').error('Cache store error', { key, error: e.message });
    }
  }
  
  static invalidate(pattern, type = this.CACHE_TYPES.SCRIPT) {
    // Google Apps Script doesn't support pattern-based invalidation
    // This is a placeholder for future enhancement
    Logger.createLogger('Cache').info(`Cache invalidation requested for pattern: ${pattern}`);
  }
  
  static getCacheService(type) {
    switch (type) {
      case this.CACHE_TYPES.USER:
        return CacheService.getUserCache();
      case this.CACHE_TYPES.DOCUMENT:
        return CacheService.getDocumentCache();
      default:
        return CacheService.getScriptCache();
    }
  }
  
  static clear(type = this.CACHE_TYPES.SCRIPT) {
    try {
      const cache = this.getCacheService(type);
      cache.removeAll();
      Logger.createLogger('Cache').info(`Cache cleared for type: ${type}`);
    } catch (e) {
      Logger.createLogger('Cache').error('Cache clear error', { type, error: e.message });
    }
  }
  
  static getStats() {
    return {
      ...this.stats,
      hitRate: this.stats.hits + this.stats.misses > 0 
        ? (this.stats.hits / (this.stats.hits + this.stats.misses) * 100).toFixed(2) + '%'
        : '0%'
    };
  }
}
/**
 * ===================================================================
 * CACHE SERVICE
 * ===================================================================
 * High-performance caching layer for expensive operations
 */

var Cache = {
  CACHE_TYPES: {
    SCRIPT: 'script',
    USER: 'user',
    DOCUMENT: 'document'
  },
  
  stats: {
    hits: 0,
    misses: 0
  },
  
  get: function(key, fetcher, options) {
    options = options || {};
    var ttl = options.ttl || 300;
    var type = options.type || this.CACHE_TYPES.SCRIPT;
    var force = options.force || false;
    
    if (!force) {
      var cached = this.retrieve(key, type);
      if (cached !== null) {
        this.stats.hits++;
        createLogger('Cache').debug('Cache hit for key: ' + key);
        return cached;
      }
    }
    
    this.stats.misses++;
    createLogger('Cache').debug('Cache miss for key: ' + key);
    var value = fetcher();
    this.store(key, value, ttl, type);
    return value;
  },
  
  retrieve: function(key, type) {
    try {
      var cache = this.getCacheService(type);
      var cached = cache.get(key);
      return cached ? JSON.parse(cached) : null;
    } catch (e) {
      createLogger('Cache').error('Cache retrieve error', { key: key, error: e.message });
      return null;
    }
  },
  
  store: function(key, value, ttl, type) {
    try {
      var cache = this.getCacheService(type);
      cache.put(key, JSON.stringify(value), ttl);
    } catch (e) {
      createLogger('Cache').error('Cache store error', { key: key, error: e.message });
    }
  },
  
  invalidate: function(pattern, type) {
    type = type || this.CACHE_TYPES.SCRIPT;
    // Google Apps Script doesn't support pattern-based invalidation
    // This is a placeholder for future enhancement
    createLogger('Cache').info('Cache invalidation requested for pattern: ' + pattern);
  },
  
  getCacheService: function(type) {
    switch (type) {
      case this.CACHE_TYPES.USER:
        return CacheService.getUserCache();
      case this.CACHE_TYPES.DOCUMENT:
        return CacheService.getDocumentCache();
      default:
        return CacheService.getScriptCache();
    }
  },
  
  clear: function(type) {
    type = type || this.CACHE_TYPES.SCRIPT;
    try {
      var cache = this.getCacheService(type);
      cache.removeAll();
      createLogger('Cache').info('Cache cleared for type: ' + type);
    } catch (e) {
      createLogger('Cache').error('Cache clear error', { type: type, error: e.message });
    }
  },
  
  getStats: function() {
    var total = this.stats.hits + this.stats.misses;
    var hitRate = total > 0 ? (this.stats.hits / total * 100).toFixed(2) + '%' : '0%';
    
    return {
      hits: this.stats.hits,
      misses: this.stats.misses,
      hitRate: hitRate
    };
  }
};
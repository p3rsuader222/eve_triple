(function (global) {
    function normalizeText(value) {
        return String(value || '').replace(/\s+/g, ' ').trim();
    }

    function getStoredLanguage(storageKey) {
        try {
            return global.localStorage && global.localStorage.getItem(storageKey);
        } catch (error) {
            return null;
        }
    }

    function setStoredLanguage(storageKey, lang) {
        try {
            if (global.localStorage) global.localStorage.setItem(storageKey, lang);
        } catch (error) {
            // Ignore restricted storage.
        }
    }

    function detectInitialLanguage(options) {
        const params = new URLSearchParams(global.location.search);
        return params.get('lang') ||
            getStoredLanguage(options.storageKey) ||
            options.defaultLanguage ||
            'en';
    }

    function createEveI18n(options) {
        const config = Object.assign({
            defaultLanguage: 'en',
            storageKey: 'eve_language',
            languages: ['en', 'lv', 'et'],
            translations: {}
        }, options || {});
        const nodeOriginals = new WeakMap();
        const attrOriginals = new WeakMap();
        let currentLanguage = normalizeLanguage(detectInitialLanguage(config));

        function normalizeLanguage(lang) {
            return config.languages.includes(lang) ? lang : config.defaultLanguage;
        }

        function dictionary(lang) {
            return config.translations[normalizeLanguage(lang)] || {};
        }

        function translate(value, lang) {
            const source = normalizeText(value);
            if (!source) return value;
            if (normalizeLanguage(lang || currentLanguage) === config.defaultLanguage) return source;
            return dictionary(lang || currentLanguage)[source] || source;
        }

        function translateNode(node) {
            const original = nodeOriginals.get(node) || node.nodeValue;
            if (!nodeOriginals.has(node)) nodeOriginals.set(node, original);
            const leading = String(original).match(/^\s*/)[0];
            const trailing = String(original).match(/\s*$/)[0];
            const translated = translate(original);
            node.nodeValue = leading + translated + trailing;
        }

        function shouldSkipTextNode(node) {
            const parent = node.parentElement;
            if (!parent) return true;
            if (!normalizeText(node.nodeValue)) return true;
            return !!parent.closest('script, style, code, pre, textarea, [data-i18n-skip]');
        }

        function translateAttribute(el, attr) {
            if (!el.hasAttribute(attr)) return;
            let originals = attrOriginals.get(el);
            if (!originals) {
                originals = {};
                attrOriginals.set(el, originals);
            }
            if (!Object.prototype.hasOwnProperty.call(originals, attr)) {
                originals[attr] = el.getAttribute(attr);
            }
            el.setAttribute(attr, translate(originals[attr]));
        }

        function apply(root) {
            const base = root || global.document;
            if (global.document && global.document.documentElement) {
                global.document.documentElement.lang = currentLanguage;
            }

            const walker = global.document.createTreeWalker(
                base.body || base,
                NodeFilter.SHOW_TEXT,
                {
                    acceptNode: function (node) {
                        return shouldSkipTextNode(node)
                            ? NodeFilter.FILTER_REJECT
                            : NodeFilter.FILTER_ACCEPT;
                    }
                }
            );

            const nodes = [];
            while (walker.nextNode()) nodes.push(walker.currentNode);
            nodes.forEach(translateNode);

            (base.querySelectorAll ? base : global.document).querySelectorAll('[placeholder], [title], [aria-label], [alt]').forEach(function (el) {
                translateAttribute(el, 'placeholder');
                translateAttribute(el, 'title');
                translateAttribute(el, 'aria-label');
                translateAttribute(el, 'alt');
            });
        }

        function setLanguage(lang, options) {
            currentLanguage = normalizeLanguage(lang);
            if (!options || options.persist !== false) {
                setStoredLanguage(config.storageKey, currentLanguage);
            }
            apply();
            if (typeof config.onChange === 'function') {
                config.onChange(currentLanguage);
            }
            global.dispatchEvent(new CustomEvent('eve:language-change', { detail: { lang: currentLanguage } }));
        }

        global.addEventListener('message', function (event) {
            const data = event.data || {};
            if (data && data.type === 'eve:setLanguage') {
                setLanguage(data.lang, { persist: false });
            }
        });

        return {
            apply: apply,
            setLanguage: setLanguage,
            t: translate,
            getLanguage: function () { return currentLanguage; },
            languages: config.languages.slice()
        };
    }

    global.createEveI18n = createEveI18n;
}(window));

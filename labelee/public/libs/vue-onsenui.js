/* vue-onsenui v2.6.1 - 2018-06-08 */

(function (global, factory) {
	typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory(require('onsenui')) :
	typeof define === 'function' && define.amd ? define(['onsenui'], factory) :
	(global.VueOnsen = factory(global.ons));
}(this, (function (ons) { 'use strict';

ons = ons && ons.hasOwnProperty('default') ? ons['default'] : ons;

var setup = function (ons$$1) {
  return Object.keys(ons$$1).filter(function (k) {
    return [/^is/, /^disable/, /^enable/, /^mock/, /^open/, /^set/, /animit/, /elements/, /fastClick/, /GestureDetector/, /notification/, /orientation/, /platform/, /ready/].some(function (t) {
      return k.match(t);
    });
  }).reduce(function (r, k) {
    r[k] = ons$$1[k];
    return r;
  }, { _ons: ons$$1 });
};

var capitalize = function capitalize(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
};

var camelize = function camelize(string) {
  return string.toLowerCase().replace(/-([a-z])/g, function (m, l) {
    return l.toUpperCase();
  });
};

var eventToHandler = function eventToHandler(name) {
  return '_on' + capitalize(name);
};

var handlerToProp = function handlerToProp(name) {
  return name.slice(2).charAt(0).toLowerCase() + name.slice(2).slice(1);
};

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) {
  return typeof obj;
} : function (obj) {
  return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj;
};



















var defineProperty = function (obj, key, value) {
  if (key in obj) {
    Object.defineProperty(obj, key, {
      value: value,
      enumerable: true,
      configurable: true,
      writable: true
    });
  } else {
    obj[key] = value;
  }

  return obj;
};

var _extends = Object.assign || function (target) {
  for (var i = 1; i < arguments.length; i++) {
    var source = arguments[i];

    for (var key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        target[key] = source[key];
      }
    }
  }

  return target;
};



































var toConsumableArray = function (arr) {
  if (Array.isArray(arr)) {
    for (var i = 0, arr2 = Array(arr.length); i < arr.length; i++) arr2[i] = arr[i];

    return arr2;
  } else {
    return Array.from(arr);
  }
};

/* Private */
var _setupDBB = function _setupDBB(component) {
  var dbb = 'onDeviceBackButton';
  // Call original handler or parent handler by default
  var handler = component[dbb] || component.$el[dbb] && component.$el[dbb]._callback || function (e) {
    return e.callParentHandler();
  };

  component.$el[dbb] = function (event) {
    var runDefault = true;

    component.$emit(handlerToProp(dbb), _extends({}, event, {
      preventDefault: function preventDefault() {
        return runDefault = false;
      }
    }));

    runDefault && handler(event);
  };

  component._isDBBSetup = true;
};

/* Public */
// Device Back Button Handler
var deriveDBB = {
  mounted: function mounted() {
    _setupDBB(this);
  },


  // Core destroys deviceBackButton handlers on disconnectedCallback.
  // This fixes the behavior for <keep-alive> component.
  activated: function activated() {
    this._isDBBSetup === false && _setupDBB(this);
  },
  deactivated: function deactivated() {
    this._isDBBSetup === true && (this._isDBBSetup = false);
  },
  destroyed: function destroyed() {
    this.$el.onDeviceBackButton && this.$el.onDeviceBackButton.destroy();
  }
};

var deriveEvents = {
  computed: {
    unrecognizedListeners: function unrecognizedListeners() {
      var _this = this;

      var name = camelize('-' + this.$options._componentTag.slice(6));
      return Object.keys(this.$listeners || {}).filter(function (k) {
        return (_this.$ons.elements[name].events || []).indexOf(k) === -1;
      }).reduce(function (r, k) {
        r[k] = _this.$listeners[k];
        return r;
      }, {});
    }
  },

  mounted: function mounted() {
    var _this2 = this;

    this._handlers = {};

    (this.$el.constructor.events || []).forEach(function (key) {
      _this2._handlers[eventToHandler(key)] = function (event) {
        // Filter events from different components with the same name
        if (event.target === _this2.$el || !/^ons-/i.test(event.target.tagName)) {
          _this2.$emit(key, event);
        }
      };
      _this2.$el.addEventListener(key, _this2._handlers[eventToHandler(key)]);
    });
  },
  beforeDestroy: function beforeDestroy() {
    var _this3 = this;

    Object.keys(this._handlers).forEach(function (key) {
      _this3.$el.removeEventListener(key, _this3._handlers[key]);
    });
    this._handlers = null;
  }
};

/* Private */
var _toggleVisibility = function _toggleVisibility() {
  if (typeof this.visible === 'boolean' && this.visible !== this.$el.visible) {
    this.$el[this.visible ? 'show' : 'hide'].call(this.$el, this.normalizedOptions || this.options);
  }
};
var _teleport = function _teleport() {
  if (!this._isDestroyed && (!this.$el.parentNode || this.$el.parentNode !== document.body)) {
    document.body.appendChild(this.$el);
  }
};
var _unmount = function _unmount() {
  var _this = this;

  if (this.$el.visible === true) {
    this.$el.hide().then(function () {
      return _this.$el.remove();
    });
  } else {
    this.$el.remove();
  }
};

/* Public */
// Components that can be shown or hidden
var hidable = {
  props: {
    visible: {
      type: Boolean,
      default: undefined // Avoid casting to false
    }
  },

  watch: {
    visible: function visible() {
      _toggleVisibility.call(this);
    }
  },

  mounted: function mounted() {
    var _this2 = this;

    this.$nextTick(function () {
      return _toggleVisibility.call(_this2);
    });
  },
  activated: function activated() {
    var _this3 = this;

    this.$nextTick(function () {
      return _toggleVisibility.call(_this3);
    });
  }
};

// Components with 'options' property
var hasOptions = {
  props: {
    options: {
      type: Object,
      default: function _default() {
        return {};
      }
    }
  }
};

// Provides itself to its descendants
var selfProvider = {
  provide: function provide() {
    return defineProperty({}, this.$options._componentTag.slice(6), this);
  }
};

// Common event for Dialogs
var dialogCancel = {
  mounted: function mounted() {
    var _this4 = this;

    this.$on('dialog-cancel', function () {
      return _this4.$emit('update:visible', false);
    });
  }
};

// Moves the element to a global position
var portal = {
  mounted: function mounted() {
    _teleport.call(this);
  },
  updated: function updated() {
    _teleport.call(this);
  },
  activated: function activated() {
    _teleport.call(this);
  },
  deactivated: function deactivated() {
    _unmount.call(this);
  },
  beforeDestroy: function beforeDestroy() {
    _unmount.call(this);
  }
};

var modifier = {
  props: {
    modifier: {
      type: [String, Array, Object]
    }
  },

  computed: {
    normalizedModifier: function normalizedModifier() {
      var modifier = this.modifier;

      if (typeof modifier === 'string') {
        return modifier;
      }

      if (Array.isArray(modifier)) {
        return modifier.join(' ');
      }

      if ((typeof modifier === 'undefined' ? 'undefined' : _typeof(modifier)) === 'object') {
        return Object.keys(modifier).reduce(function (acc, key) {
          return acc + (modifier[key] ? ' ' + key : '');
        }, '').trim();
      }

      return false;
    }
  }
};

var _props;
var _props2;

/* Private */
var model = {
  prop: 'modelProp',
  event: 'modelEvent'
};

/* Public */

// Generic input
var modelInput = {
  model: model,
  props: (_props = {}, defineProperty(_props, model.prop, [Number, String]), defineProperty(_props, model.event, {
    type: String,
    default: 'input'
  }), _props),

  methods: {
    _updateValue: function _updateValue() {
      if (this[model.prop] !== undefined && this.$el.value !== this[model.prop]) {
        this.$el.value = this[model.prop];
      }
    },
    _onModelEvent: function _onModelEvent(event) {
      this.$emit(model.event, event.target.value);
    }
  },

  watch: defineProperty({}, model.prop, function () {
    this._updateValue();
  }),

  mounted: function mounted() {
    this._updateValue();
    this.$el.addEventListener(this[model.event], this._onModelEvent);
  },
  beforeDestroy: function beforeDestroy() {
    this.$el.removeEventListener(this[model.event], this._onModelEvent);
  }
};

// Checkable inputs
var modelCheckbox = {
  mixins: [modelInput],

  props: (_props2 = {}, defineProperty(_props2, model.prop, [Array, Boolean]), defineProperty(_props2, model.event, {
    type: String,
    default: 'change'
  }), _props2),

  methods: {
    _updateValue: function _updateValue() {
      if (this[model.prop] instanceof Array) {
        this.$el.checked = this[model.prop].indexOf(this.$el.value) >= 0;
      } else {
        this.$el.checked = this[model.prop];
      }
    },
    _onModelEvent: function _onModelEvent(event) {
      var _event$target = event.target,
          value = _event$target.value,
          checked = _event$target.checked;

      var newValue = void 0;

      if (this[model.prop] instanceof Array) {
        // Is Array
        var index = this[model.prop].indexOf(value);
        var included = index >= 0;

        if (included && !checked) {
          newValue = [].concat(toConsumableArray(this[model.prop].slice(0, index)), toConsumableArray(this[model.prop].slice(index + 1, this[model.prop].length)));
        }

        if (!included && checked) {
          newValue = [].concat(toConsumableArray(this[model.prop]), [value]);
        }
      } else {
        // Is Boolean
        newValue = checked;
      }

      // Emit if value changed
      newValue !== undefined && this.$emit(model.event, newValue);
    }
  }
};

// Radio input
var modelRadio = {
  mixins: [modelInput],
  props: defineProperty({}, model.event, {
    type: String,
    default: 'change'
  }),

  methods: {
    _updateValue: function _updateValue() {
      this.$el.checked = this[model.prop] === this.$el.value;
    },
    _onModelEvent: function _onModelEvent(event) {
      var _event$target2 = event.target,
          value = _event$target2.value,
          checked = _event$target2.checked;

      checked && this.$emit(model.event, value);
    }
  }
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-toolbar';
var VOnsToolbar = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-toolbar', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-toolbar',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-bottom-toolbar';
var VOnsBottomToolbar = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-bottom-toolbar', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-bottom-toolbar',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-toolbar-button';
var VOnsToolbarButton = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-toolbar-button', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-toolbar-button',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-alert-dialog-button';
var VOnsAlertDialogButton = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-alert-dialog-button', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-alert-dialog-button',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-button';
var VOnsButton = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-button', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-button',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-icon';
var VOnsIcon = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-icon', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-icon',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-card';
var VOnsCard = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-card', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-card',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-list';
var VOnsList = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-list', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-list',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-list-item';
var VOnsListItem = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-list-item', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-list-item',
  mixins: [deriveEvents, modifier],
  props: {
    expanded: {
      type: Boolean
    }
  },

  watch: {
    expanded: function expanded() {
      var action = this.expanded ? 'show' : 'hide';
      this.$el[action + 'Expansion']();
    }
  }
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-list-title';
var VOnsListTitle = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-list-title', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-list-title',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-list-header';
var VOnsListHeader = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-list-header', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-list-header',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-ripple';
var VOnsRipple = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-ripple', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-ripple',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-row';
var VOnsRow = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-row', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-row',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-col';
var VOnsCol = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-col', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-col',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-progress-bar';
var VOnsProgressBar = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-progress-bar', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-progress-bar',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-progress-circular';
var VOnsProgressCircular = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-progress-circular', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-progress-circular',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-carousel-item';
var VOnsCarouselItem = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-carousel-item', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-carousel-item',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-splitter-mask';
var VOnsSplitterMask = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-splitter-mask', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-splitter-mask',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-splitter-content';
var VOnsSplitterContent = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-splitter-content', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-splitter-content',
  mixins: [deriveEvents]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-splitter';
var VOnsSplitter = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-splitter', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-splitter',
  mixins: [deriveEvents, selfProvider, deriveDBB]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-switch';
var VOnsSwitch = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-switch', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-switch',
  mixins: [deriveEvents, modelCheckbox, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-checkbox';
var VOnsCheckbox = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-checkbox', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-checkbox',
  mixins: [deriveEvents, modelCheckbox, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-input';
var VOnsInput = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-input', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-input',
  mixins: [deriveEvents, modelInput, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-search-input';
var VOnsSearchInput = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-search-input', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-search-input',
  mixins: [deriveEvents, modelInput, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-range';
var VOnsRange = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-range', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-range',
  mixins: [deriveEvents, modelInput, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-radio';
var VOnsRadio = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-radio', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-radio',
  mixins: [deriveEvents, modelRadio, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-fab';
var VOnsFab = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-fab', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-fab',
  mixins: [deriveEvents, hidable, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-speed-dial-item';
var VOnsSpeedDialItem = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-speed-dial-item', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-speed-dial-item',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-dialog';
var VOnsDialog = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-dialog', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-dialog',
  mixins: [deriveEvents, hidable, hasOptions, dialogCancel, deriveDBB, portal, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-action-sheet';
var VOnsActionSheet = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-action-sheet', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-action-sheet',
  mixins: [deriveEvents, hidable, hasOptions, dialogCancel, deriveDBB, portal, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-action-sheet-button';
var VOnsActionSheetButton = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-action-sheet-button', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-action-sheet-button',
  mixins: [deriveEvents, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-modal';
var VOnsModal = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-modal', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-modal',
  mixins: [deriveEvents, hidable, hasOptions, deriveDBB, portal, modifier]
};

/* This file was generated automatically by 'generate-components' task in bindings/vue/gulpfile.babel.js */
// 'ons-toast';
var VOnsToast = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-toast', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-toast',
  mixins: [deriveEvents, hidable, hasOptions, deriveDBB, portal, modifier]
};

// 'ons-popover';
var VOnsPopover = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-popover', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-popover',
  mixins: [hidable, hasOptions, dialogCancel, deriveEvents, deriveDBB, portal],

  props: {
    target: {
      validator: function validator(value) {
        return value._isVue || typeof value === 'string' || value instanceof Event || value instanceof HTMLElement;
      }
    }
  },

  computed: {
    normalizedTarget: function normalizedTarget() {
      if (this.target && this.target._isVue) {
        return this.target.$el;
      }
      return this.target;
    },
    normalizedOptions: function normalizedOptions() {
      if (this.target) {
        return _extends({
          target: this.normalizedTarget
        }, this.options);
      }
      return this.options;
    }
  }
};

// 'ons-alert-dialog';
var VOnsAlertDialog = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-alert-dialog', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.unrecognizedListeners), [_c('div', { staticClass: "alert-dialog-title" }, [_vm._t("title", [_vm._v(_vm._s(_vm.title))])], 2), _vm._v(" "), _c('div', { staticClass: "alert-dialog-content" }, [_vm._t("default")], 2), _vm._v(" "), _c('div', { staticClass: "alert-dialog-footer" }, [_vm._t("footer", _vm._l(_vm.footer, function (handler, key) {
      return _c('ons-alert-dialog-button', { key: key, on: { "click": handler } }, [_vm._v(_vm._s(key))]);
    }))], 2)]);
  }, staticRenderFns: [],
  name: 'v-ons-alert-dialog',
  mixins: [hidable, hasOptions, dialogCancel, deriveEvents, deriveDBB, portal, modifier],

  props: {
    title: {
      type: String
    },
    footer: {
      type: Object,
      validator: function validator(value) {
        return Object.keys(value).every(function (key) {
          return value[key] instanceof Function;
        });
      }
    }
  }
};

// 'ons-speed-dial';
var VOnsSpeedDial = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-speed-dial', { domProps: { "onClick": _vm.action } }, [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-speed-dial',
  mixins: [deriveEvents, hidable],

  props: {
    open: {
      type: Boolean,
      default: undefined
    }
  },

  methods: {
    action: function action() {
      var runDefault = true;
      this.$emit('click', { preventDefault: function preventDefault() {
          return runDefault = false;
        } });

      if (runDefault) {
        this.$el.toggleItems();
      }
    },
    _shouldUpdate: function _shouldUpdate() {
      return this.open !== undefined && this.open !== this.$el.isOpen();
    },
    _updateToggle: function _updateToggle() {
      this._shouldUpdate() && this.$el[this.open ? 'showItems' : 'hideItems'].call(this.$el);
    }
  },

  watch: {
    open: function open() {
      this._updateToggle();
    }
  },

  mounted: function mounted() {
    var _this = this;

    this.$on(['open', 'close'], function () {
      return _this._shouldUpdate() && _this.$emit('update:open', _this.$el.isOpen());
    });

    this._updateToggle();
  }
};

// 'ons-carousel';
var VOnsCarousel = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-carousel', _vm._g({ attrs: { "initial-index": _vm.index }, domProps: { "onSwipe": _vm.onSwipe }, on: { "postchange": function postchange($event) {
          if ($event.target !== $event.currentTarget) {
            return null;
          }_vm.$emit('update:index', $event.activeIndex);
        } } }, _vm.unrecognizedListeners), [_c('div', [_vm._t("default")], 2), _vm._v(" "), _c('div')]);
  }, staticRenderFns: [],
  name: 'v-ons-carousel',
  mixins: [hasOptions, deriveEvents],

  props: {
    index: {
      type: Number
    },
    onSwipe: {
      type: Function
    }
  },

  watch: {
    index: function index() {
      if (this.index !== this.$el.getActiveIndex()) {
        this.$el.setActiveIndex(this.index, this.options);
      }
    }
  }
};

// 'ons-tab';

var VOnsTab = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-tab', { attrs: { "active": _vm.active }, domProps: { "onClick": _vm.action } });
  }, staticRenderFns: [],
  name: 'v-ons-tab',
  inject: ['tabbar'],

  props: {
    page: {},
    props: {},
    active: {
      type: Boolean
    }
  },

  methods: {
    action: function action() {
      var runDefault = true;
      this.$emit('click', { preventDefault: function preventDefault() {
          return runDefault = false;
        } });

      if (runDefault) {
        this.tabbar.$el.setActiveTab(this.$el.index, _extends({ reject: false }, this.tabbar.options));
      }
    }
  },

  watch: {
    active: function active() {
      this.$el.setActive(this.active);
    }
  }
};

// 'ons-tabbar';
var VOnsTabbar = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-tabbar', _vm._g({ attrs: { "activeIndex": _vm.index, "modifier": _vm.normalizedModifier }, domProps: { "onSwipe": _vm.onSwipe }, on: { "prechange": function prechange($event) {
          if ($event.target !== $event.currentTarget) {
            return null;
          }_vm.$nextTick(function () {
            return !$event.detail.canceled && _vm.$emit('update:index', $event.index);
          });
        } } }, _vm.unrecognizedListeners), [_c('div', { staticClass: "tabbar__content" }, [_c('div', [_vm._t("pages", _vm._l(_vm.tabs, function (tab) {
      return _c(tab.page, _vm._g(_vm._b({ key: tab.page.key || tab.page.name || _vm._tabKey(tab), tag: "component" }, 'component', tab.props, false), _vm.unrecognizedListeners));
    }))], 2), _vm._v(" "), _c('div')]), _vm._v(" "), _c('div', { staticClass: "tabbar", style: _vm.tabbarStyle }, [_vm._t("default", _vm._l(_vm.tabs, function (tab) {
      return _c('v-ons-tab', _vm._b({ key: _vm._tabKey(tab) }, 'v-ons-tab', tab, false));
    })), _vm._v(" "), _c('div', { staticClass: "tabbar__border" })], 2)]);
  }, staticRenderFns: [],
  name: 'v-ons-tabbar',
  mixins: [deriveEvents, hasOptions, hidable, selfProvider, modifier],

  props: {
    index: {
      type: Number
    },
    tabs: {
      type: Array,
      validator: function validator(value) {
        return value.every(function (tab) {
          return ['icon', 'label', 'page'].some(function (prop) {
            return !!Object.getOwnPropertyDescriptor(tab, prop);
          });
        });
      }
    },
    onSwipe: {
      type: Function
    },
    tabbarStyle: {
      type: null
    }
  },

  methods: {
    _tabKey: function _tabKey(tab) {
      return tab.key || tab.label || tab.icon;
    }
  },

  watch: {
    index: function index() {
      if (this.index !== this.$el.getActiveTabIndex()) {
        this.$el.setActiveTab(this.index, _extends({ reject: false }, this.options));
      }
    }
  }
};

// 'ons-back-button';
var VOnsBackButton = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-back-button', { domProps: { "onClick": _vm.action } }, [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-back-button',
  inject: ['navigator'],
  mixins: [modifier],

  methods: {
    action: function action() {
      var runDefault = true;
      this.$emit('click', { preventDefault: function preventDefault() {
          return runDefault = false;
        } });

      if (runDefault && this.navigator.pageStack.length > 1) {
        this.navigator.popPage();
      }
    }
  }
};

// 'ons-navigator';
var VOnsNavigator = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-navigator', _vm._g({ domProps: { "options": _vm.options }, on: { "postpop": function postpop($event) {
          if ($event.target !== $event.currentTarget) {
            return null;
          }_vm._checkSwipe($event);
        } } }, _vm.unrecognizedListeners), [_vm._t("default", _vm._l(_vm.pageStack, function (page) {
      return _c(page, _vm._g(_vm._b({ key: page.key || page.name, tag: "component" }, 'component', page.onsNavigatorProps, false), _vm.unrecognizedListeners));
    }))], 2);
  }, staticRenderFns: [],
  name: 'v-ons-navigator',
  mixins: [hasOptions, selfProvider, deriveEvents, deriveDBB],

  props: {
    pageStack: {
      type: Array,
      required: true
    },
    popPage: {
      type: Function,
      default: function _default() {
        this.pageStack.pop();
      }
    }
  },

  methods: {
    isReady: function isReady() {
      if (this.hasOwnProperty('_ready') && this._ready instanceof Promise) {
        return this._ready;
      }
      return Promise.resolve();
    },
    onDeviceBackButton: function onDeviceBackButton(event) {
      if (this.pageStack.length > 1) {
        this.popPage();
      } else {
        event.callParentHandler();
      }
    },
    _findScrollPage: function _findScrollPage(page) {
      var nextPage = page._contentElement.children.length === 1 && this.$ons._ons._util.getTopPage(page._contentElement.children[0]);
      return nextPage ? this._findScrollPage(nextPage) : page;
    },
    _eachPage: function _eachPage(start, end, cb) {
      for (var i = start; i < end; i++) {
        cb(this.$children[i].$el);
      }
    },
    _reattachPage: function _reattachPage(pageElement) {
      var position = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
      var restoreScroll = arguments[2];

      this.$el.insertBefore(pageElement, position);
      restoreScroll instanceof Function && restoreScroll();
      pageElement._isShown = true;
    },
    _redetachPage: function _redetachPage(pageElement) {
      pageElement._destroy();
      return Promise.resolve();
    },
    _animate: function _animate(_ref) {
      var _this = this;

      var lastLength = _ref.lastLength,
          currentLength = _ref.currentLength,
          lastTopPage = _ref.lastTopPage,
          currentTopPage = _ref.currentTopPage,
          restoreScroll = _ref.restoreScroll;

      var pushedOptions = this.pageStack[this.pageStack.length - 1].onsNavigatorOptions || currentTopPage.__vue__.onsNavigatorOptions || {};

      // Push
      if (currentLength > lastLength) {
        var isReattached = false;
        if (lastTopPage.parentElement !== this.$el) {
          this._reattachPage(lastTopPage, this.$el.children[lastLength - 1], restoreScroll);
          isReattached = true;
          lastLength--;
        }

        this._eachPage(lastLength, currentLength, function (el) {
          el.style.visibility = 'hidden';
        });
        this._eachPage(lastLength, currentLength - 1, function (el) {
          el.pushedOptions = pushedOptions;
        });

        return this.$el._pushPage(_extends({}, pushedOptions, { leavePage: lastTopPage })).then(function () {
          setImmediate(function () {
            _this._eachPage(lastLength, currentLength, function (el) {
              el.style.visibility = '';
            });
            _this._eachPage(lastLength - 1, currentLength - 1, function (el) {
              el.style.display = 'none';
            });
          });

          if (isReattached) {
            _this._redetachPage(lastTopPage);
          }
        });
      }

      // Pop
      if (currentLength < lastLength) {
        this._reattachPage(lastTopPage, null, restoreScroll);
        return this.$el._popPage({}, function () {
          return _this._redetachPage(lastTopPage);
        });
      }

      // Replace page
      currentTopPage.style.visibility = 'hidden';
      this._reattachPage(lastTopPage, currentTopPage, restoreScroll);
      return this.$el._pushPage(_extends({}, pushedOptions, { _replacePage: true })).then(function () {
        return _this._redetachPage(lastTopPage);
      });
    },
    _checkSwipe: function _checkSwipe(event) {
      if (this.$el.hasAttribute('swipeable') && event.leavePage !== this.$el.lastChild && event.leavePage === this.$children[this.$children.length - 1].$el) {
        this.popPage();
      }
    }
  },

  watch: {
    pageStack: function pageStack(after, before) {
      if (this.$el.hasAttribute('swipeable') && this.$children.length !== this.$el.children.length) {
        return;
      }

      var propWasMutated = after === before; // Can be mutated or replaced
      var lastTopPage = this.$children[this.$children.length - 1].$el;
      var scrollElement = this._findScrollPage(lastTopPage);
      var scrollValue = scrollElement.scrollTop || 0;

      this._pageStackUpdate = {
        lastTopPage: lastTopPage,
        lastLength: propWasMutated ? this.$children.length : before.length,
        currentLength: !propWasMutated && after.length,
        restoreScroll: function restoreScroll() {
          return scrollElement.scrollTop = scrollValue;
        }
      };

      // this.$nextTick(() => { }); // Waits too long, updated() hook is faster and prevents flickerings
    }
  },

  updated: function updated() {
    if (this._pageStackUpdate) {
      var currentTopPage = this.$children[this.$children.length - 1].$el;
      var _pageStackUpdate = this._pageStackUpdate,
          lastTopPage = _pageStackUpdate.lastTopPage,
          currentLength = _pageStackUpdate.currentLength;
      var _pageStackUpdate2 = this._pageStackUpdate,
          lastLength = _pageStackUpdate2.lastLength,
          restoreScroll = _pageStackUpdate2.restoreScroll;

      currentLength = currentLength === false ? this.$children.length : currentLength;

      if (currentTopPage !== lastTopPage) {
        this._ready = this._animate({ lastLength: lastLength, currentLength: currentLength, lastTopPage: lastTopPage, currentTopPage: currentTopPage, restoreScroll: restoreScroll });
      } else if (currentLength !== lastLength) {
        currentTopPage.updateBackButton(currentLength > 1);
      }

      lastTopPage = currentTopPage = this._pageStackUpdate = null;
    }
  }
};

// 'ons-splitter-side';
var VOnsSplitterSide = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-splitter-side', _vm._g({}, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-splitter-side',
  mixins: [hasOptions, deriveEvents],

  props: {
    open: {
      type: Boolean,
      default: undefined
    }
  },

  methods: {
    action: function action() {
      this._shouldUpdate() && this.$el[this.open ? 'open' : 'close'].call(this.$el, this.options).catch(function () {});
    },
    _shouldUpdate: function _shouldUpdate() {
      return this.open !== undefined && this.open !== this.$el.isOpen;
    }
  },

  watch: {
    open: function open() {
      this.action();
    }
  },

  mounted: function mounted() {
    var _this = this;

    this.$on(['postopen', 'postclose', 'modechange'], function () {
      return _this._shouldUpdate() && _this.$emit('update:open', _this.$el.isOpen);
    });

    this.action();
  }
};

// 'ons-lazy-repeat';

var VOnsLazyRepeat = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-lazy-repeat');
  }, staticRenderFns: [],
  name: 'v-ons-lazy-repeat',

  props: {
    renderItem: {
      type: Function,
      required: true,
      validator: function validator(value) {
        var component = value(0);
        if (component._isVue && !component._isMounted) {
          component.$destroy();
          return true;
        }
        return false;
      }
    },
    length: {
      type: Number,
      required: true
    },
    calculateItemHeight: {
      type: Function,
      default: undefined
    }
  },

  data: function data() {
    return {
      provider: null
    };
  },


  methods: {
    _setup: function _setup() {
      var _this = this;

      this.provider && this.provider.destroy();

      var delegate = new this.$ons._ons._internal.LazyRepeatDelegate({
        calculateItemHeight: this.calculateItemHeight,
        createItemContent: function createItemContent(i) {
          return _this.renderItem(i).$mount().$el;
        },
        destroyItem: function destroyItem(i, _ref) {
          var element = _ref.element;
          return element.__vue__.$destroy();
        },
        countItems: function countItems() {
          return _this.length;
        }
      }, null);

      this.provider = new this.$ons._ons._internal.LazyRepeatProvider(this.$parent.$el, delegate);
    },
    refresh: function refresh() {
      return this.provider.refresh();
    }
  },

  watch: {
    renderItem: function renderItem() {
      this._setup();
    },
    length: function length() {
      this._setup();
    },
    calculateItemHeight: function calculateItemHeight() {
      this._setup();
    }
  },

  mounted: function mounted() {
    this._setup();
    this.$vnode.context.$on('refresh', this.refresh);
  },
  beforeDestroy: function beforeDestroy() {
    this.$vnode.context.$off('refresh', this.refresh);

    // This will destroy the provider once the rendered element
    // is detached (detachedCallback). Therefore, animations
    // have time to finish before elements start to disappear.
    // It cannot be set earlier in order to prevent accidental
    // destroys if this element is retached by something else.
    this.$el._lazyRepeatProvider = this.provider;
    this.provider = null;
  }
};

// 'ons-select';
var VOnsSelect = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-select', _vm._g({ attrs: { "modifier": _vm.normalizedModifier } }, _vm.$listeners), [_c('select', [_vm._t("default")], 2)]);
  }, staticRenderFns: [],
  name: 'v-ons-select',
  mixins: [modelInput, modifier]
};

// 'ons-segment';
var VOnsSegment = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-segment', { attrs: { "active-index": _vm.index }, on: { "postchange": function postchange($event) {
          if ($event.target !== $event.currentTarget) {
            return null;
          }_vm.$emit('update:index', $event.index);
        } } }, [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-segment',
  mixins: [deriveEvents],

  props: {
    index: {
      type: Number
    }
  },

  watch: {
    index: function index() {
      if (this.index !== this.$el.getActiveButtonIndex()) {
        this.$el.setActiveButton(this.index, { reject: false });
      }
    }
  }
};

// 'ons-pull-hook';
var VOnsPullHook = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-pull-hook', _vm._g({ domProps: { "onAction": _vm.action, "onPull": _vm.onPull } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-pull-hook',
  mixins: [deriveEvents],

  props: {
    action: {
      type: Function
    },
    onPull: {
      type: Function
    }
  }
};

// 'ons-page';
var VOnsPage = { render: function render() {
    var _vm = this;var _h = _vm.$createElement;var _c = _vm._self._c || _h;return _c('ons-page', _vm._g({ attrs: { "modifier": _vm.normalizedModifier }, domProps: { "onInfiniteScroll": _vm.infiniteScroll } }, _vm.unrecognizedListeners), [_vm._t("default")], 2);
  }, staticRenderFns: [],
  name: 'v-ons-page',
  mixins: [deriveEvents, deriveDBB, modifier],

  props: {
    infiniteScroll: {
      type: Function
    }
  }
};

// Generic components:


var components = Object.freeze({
	VOnsToolbar: VOnsToolbar,
	VOnsBottomToolbar: VOnsBottomToolbar,
	VOnsToolbarButton: VOnsToolbarButton,
	VOnsAlertDialogButton: VOnsAlertDialogButton,
	VOnsButton: VOnsButton,
	VOnsIcon: VOnsIcon,
	VOnsCard: VOnsCard,
	VOnsList: VOnsList,
	VOnsListItem: VOnsListItem,
	VOnsListTitle: VOnsListTitle,
	VOnsListHeader: VOnsListHeader,
	VOnsRipple: VOnsRipple,
	VOnsRow: VOnsRow,
	VOnsCol: VOnsCol,
	VOnsProgressBar: VOnsProgressBar,
	VOnsProgressCircular: VOnsProgressCircular,
	VOnsCarouselItem: VOnsCarouselItem,
	VOnsSplitterMask: VOnsSplitterMask,
	VOnsSplitterContent: VOnsSplitterContent,
	VOnsSplitter: VOnsSplitter,
	VOnsSwitch: VOnsSwitch,
	VOnsCheckbox: VOnsCheckbox,
	VOnsInput: VOnsInput,
	VOnsSearchInput: VOnsSearchInput,
	VOnsRange: VOnsRange,
	VOnsRadio: VOnsRadio,
	VOnsFab: VOnsFab,
	VOnsSpeedDialItem: VOnsSpeedDialItem,
	VOnsDialog: VOnsDialog,
	VOnsActionSheet: VOnsActionSheet,
	VOnsActionSheetButton: VOnsActionSheetButton,
	VOnsModal: VOnsModal,
	VOnsToast: VOnsToast,
	VOnsPopover: VOnsPopover,
	VOnsAlertDialog: VOnsAlertDialog,
	VOnsSpeedDial: VOnsSpeedDial,
	VOnsCarousel: VOnsCarousel,
	VOnsTab: VOnsTab,
	VOnsTabbar: VOnsTabbar,
	VOnsBackButton: VOnsBackButton,
	VOnsNavigator: VOnsNavigator,
	VOnsSplitterSide: VOnsSplitterSide,
	VOnsLazyRepeat: VOnsLazyRepeat,
	VOnsSelect: VOnsSelect,
	VOnsSegment: VOnsSegment,
	VOnsPullHook: VOnsPullHook,
	VOnsPage: VOnsPage
});

var $ons = setup(ons);

$ons.install = function (Vue) {
  Object.keys(components).forEach(function (key) {
    return Vue.component(components[key].name, components[key]);
  });

  /**
   * Expose ons object.
   */
  Vue.prototype.$ons = $ons;
};

if (typeof window !== 'undefined' && window.Vue) {
  window.Vue.use({ install: $ons.install });
}

return $ons;

})));
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidnVlLW9uc2VudWkuanMiLCJzb3VyY2VzIjpbIi4uL3NyYy9zZXR1cC5qcyIsIi4uL3NyYy9pbnRlcm5hbC91dGlsLmpzIiwiLi4vc3JjL21peGlucy9kZXJpdmUuanMiLCIuLi9zcmMvbWl4aW5zL2NvbW1vbi5qcyIsIi4uL3NyYy9taXhpbnMvbW9kZWwuanMiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zVG9vbGJhci52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zQm90dG9tVG9vbGJhci52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zVG9vbGJhckJ1dHRvbi52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zQWxlcnREaWFsb2dCdXR0b24udnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0J1dHRvbi52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zSWNvbi52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zQ2FyZC52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zTGlzdC52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zTGlzdEl0ZW0udnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0xpc3RUaXRsZS52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zTGlzdEhlYWRlci52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zUmlwcGxlLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNSb3cudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0NvbC52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zUHJvZ3Jlc3NCYXIudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1Byb2dyZXNzQ2lyY3VsYXIudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0Nhcm91c2VsSXRlbS52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zU3BsaXR0ZXJNYXNrLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNTcGxpdHRlckNvbnRlbnQudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1NwbGl0dGVyLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNTd2l0Y2gudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0NoZWNrYm94LnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNJbnB1dC52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zU2VhcmNoSW5wdXQudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1JhbmdlLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNSYWRpby52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zRmFiLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNTcGVlZERpYWxJdGVtLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNEaWFsb2cudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0FjdGlvblNoZWV0LnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNBY3Rpb25TaGVldEJ1dHRvbi52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zTW9kYWwudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1RvYXN0LnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNQb3BvdmVyLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNBbGVydERpYWxvZy52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zU3BlZWREaWFsLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNDYXJvdXNlbC52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zVGFiLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNUYWJiYXIudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc0JhY2tCdXR0b24udnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc05hdmlnYXRvci52dWUiLCIuLi9zcmMvY29tcG9uZW50cy9WT25zU3BsaXR0ZXJTaWRlLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNMYXp5UmVwZWF0LnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNTZWxlY3QudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1NlZ21lbnQudnVlIiwiLi4vc3JjL2NvbXBvbmVudHMvVk9uc1B1bGxIb29rLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL1ZPbnNQYWdlLnZ1ZSIsIi4uL3NyYy9jb21wb25lbnRzL2luZGV4LmpzIiwiLi4vc3JjL2luZGV4LnVtZC5qcyJdLCJzb3VyY2VzQ29udGVudCI6WyJleHBvcnQgZGVmYXVsdCBmdW5jdGlvbihvbnMpIHtcbiAgcmV0dXJuIE9iamVjdC5rZXlzKG9ucylcbiAgICAuZmlsdGVyKGsgPT4gW1xuICAgICAgL15pcy8sXG4gICAgICAvXmRpc2FibGUvLFxuICAgICAgL15lbmFibGUvLFxuICAgICAgL15tb2NrLyxcbiAgICAgIC9eb3Blbi8sXG4gICAgICAvXnNldC8sXG4gICAgICAvYW5pbWl0LyxcbiAgICAgIC9lbGVtZW50cy8sXG4gICAgICAvZmFzdENsaWNrLyxcbiAgICAgIC9HZXN0dXJlRGV0ZWN0b3IvLFxuICAgICAgL25vdGlmaWNhdGlvbi8sXG4gICAgICAvb3JpZW50YXRpb24vLFxuICAgICAgL3BsYXRmb3JtLyxcbiAgICAgIC9yZWFkeS8sXG4gICAgXS5zb21lKHQgPT4gay5tYXRjaCh0KSkpXG4gICAgLnJlZHVjZSgociwgaykgPT4ge1xuICAgICAgcltrXSA9IG9uc1trXTtcbiAgICAgIHJldHVybiByO1xuICAgIH0sIHsgX29uczogb25zIH0pO1xufVxuIiwiZXhwb3J0IGNvbnN0IGh5cGhlbmF0ZSA9IHN0cmluZyA9PiBzdHJpbmcucmVwbGFjZSgvKFthLXpBLVpdKShbQS1aXSkvZywgJyQxLSQyJykudG9Mb3dlckNhc2UoKTtcblxuZXhwb3J0IGNvbnN0IGNhcGl0YWxpemUgPSBzdHJpbmcgPT4gc3RyaW5nLmNoYXJBdCgwKS50b1VwcGVyQ2FzZSgpICsgc3RyaW5nLnNsaWNlKDEpO1xuXG5leHBvcnQgY29uc3QgY2FtZWxpemUgPSBzdHJpbmcgPT4gc3RyaW5nLnRvTG93ZXJDYXNlKCkucmVwbGFjZSgvLShbYS16XSkvZywgKG0sIGwpID0+IGwudG9VcHBlckNhc2UoKSk7XG5cbmV4cG9ydCBjb25zdCBldmVudFRvSGFuZGxlciA9IG5hbWUgPT4gJ19vbicgKyBjYXBpdGFsaXplKG5hbWUpO1xuXG5leHBvcnQgY29uc3QgaGFuZGxlclRvUHJvcCA9IG5hbWUgPT4gbmFtZS5zbGljZSgyKS5jaGFyQXQoMCkudG9Mb3dlckNhc2UoKSArIG5hbWUuc2xpY2UoMikuc2xpY2UoMSk7XG4iLCJpbXBvcnQgeyBjYW1lbGl6ZSwgZXZlbnRUb0hhbmRsZXIsIGhhbmRsZXJUb1Byb3AgfSBmcm9tICcuLi9pbnRlcm5hbC91dGlsJztcblxuLyogUHJpdmF0ZSAqL1xuY29uc3QgX3NldHVwREJCID0gY29tcG9uZW50ID0+IHtcbiAgY29uc3QgZGJiID0gJ29uRGV2aWNlQmFja0J1dHRvbic7XG4gIC8vIENhbGwgb3JpZ2luYWwgaGFuZGxlciBvciBwYXJlbnQgaGFuZGxlciBieSBkZWZhdWx0XG4gIGNvbnN0IGhhbmRsZXIgPSBjb21wb25lbnRbZGJiXSB8fCAoY29tcG9uZW50LiRlbFtkYmJdICYmIGNvbXBvbmVudC4kZWxbZGJiXS5fY2FsbGJhY2spIHx8IChlID0+IGUuY2FsbFBhcmVudEhhbmRsZXIoKSk7XG5cbiAgY29tcG9uZW50LiRlbFtkYmJdID0gZXZlbnQgPT4ge1xuICAgIGxldCBydW5EZWZhdWx0ID0gdHJ1ZTtcblxuICAgIGNvbXBvbmVudC4kZW1pdChoYW5kbGVyVG9Qcm9wKGRiYiksIHtcbiAgICAgIC4uLmV2ZW50LFxuICAgICAgcHJldmVudERlZmF1bHQ6ICgpID0+IHJ1bkRlZmF1bHQgPSBmYWxzZVxuICAgIH0pO1xuXG4gICAgcnVuRGVmYXVsdCAmJiBoYW5kbGVyKGV2ZW50KTtcbiAgfTtcblxuICBjb21wb25lbnQuX2lzREJCU2V0dXAgPSB0cnVlO1xufTtcblxuLyogUHVibGljICovXG4vLyBEZXZpY2UgQmFjayBCdXR0b24gSGFuZGxlclxuY29uc3QgZGVyaXZlREJCID0ge1xuICBtb3VudGVkKCkge1xuICAgIF9zZXR1cERCQih0aGlzKTtcbiAgfSxcblxuICAvLyBDb3JlIGRlc3Ryb3lzIGRldmljZUJhY2tCdXR0b24gaGFuZGxlcnMgb24gZGlzY29ubmVjdGVkQ2FsbGJhY2suXG4gIC8vIFRoaXMgZml4ZXMgdGhlIGJlaGF2aW9yIGZvciA8a2VlcC1hbGl2ZT4gY29tcG9uZW50LlxuICBhY3RpdmF0ZWQoKSB7XG4gICAgdGhpcy5faXNEQkJTZXR1cCA9PT0gZmFsc2UgJiYgX3NldHVwREJCKHRoaXMpO1xuICB9LFxuXG4gIGRlYWN0aXZhdGVkKCkge1xuICAgIHRoaXMuX2lzREJCU2V0dXAgPT09IHRydWUgJiYgKHRoaXMuX2lzREJCU2V0dXAgPSBmYWxzZSk7XG4gIH0sXG5cbiAgZGVzdHJveWVkKCkge1xuICAgIHRoaXMuJGVsLm9uRGV2aWNlQmFja0J1dHRvbiAmJiB0aGlzLiRlbC5vbkRldmljZUJhY2tCdXR0b24uZGVzdHJveSgpO1xuICB9XG59O1xuXG5jb25zdCBkZXJpdmVFdmVudHMgPSB7XG4gIGNvbXB1dGVkOiB7XG4gICAgdW5yZWNvZ25pemVkTGlzdGVuZXJzKCkge1xuICAgICAgY29uc3QgbmFtZSA9IGNhbWVsaXplKCctJyArIHRoaXMuJG9wdGlvbnMuX2NvbXBvbmVudFRhZy5zbGljZSg2KSk7XG4gICAgICByZXR1cm4gT2JqZWN0LmtleXModGhpcy4kbGlzdGVuZXJzIHx8IHt9KVxuICAgICAgICAuZmlsdGVyKGsgPT4gKHRoaXMuJG9ucy5lbGVtZW50c1tuYW1lXS5ldmVudHMgfHwgW10pLmluZGV4T2YoaykgPT09IC0xKVxuICAgICAgICAucmVkdWNlKChyLCBrKSA9PiB7XG4gICAgICAgICAgcltrXSA9IHRoaXMuJGxpc3RlbmVyc1trXTtcbiAgICAgICAgICByZXR1cm4gcjtcbiAgICAgICAgfSwge30pO1xuICAgIH1cbiAgfSxcblxuICBtb3VudGVkKCkge1xuICAgIHRoaXMuX2hhbmRsZXJzID0ge307XG5cbiAgICAodGhpcy4kZWwuY29uc3RydWN0b3IuZXZlbnRzIHx8IFtdKS5mb3JFYWNoKGtleSA9PiB7XG4gICAgICB0aGlzLl9oYW5kbGVyc1tldmVudFRvSGFuZGxlcihrZXkpXSA9IGV2ZW50ID0+IHtcbiAgICAgICAgLy8gRmlsdGVyIGV2ZW50cyBmcm9tIGRpZmZlcmVudCBjb21wb25lbnRzIHdpdGggdGhlIHNhbWUgbmFtZVxuICAgICAgICBpZiAoZXZlbnQudGFyZ2V0ID09PSB0aGlzLiRlbCB8fCAhL15vbnMtL2kudGVzdChldmVudC50YXJnZXQudGFnTmFtZSkpIHtcbiAgICAgICAgICB0aGlzLiRlbWl0KGtleSwgZXZlbnQpO1xuICAgICAgICB9XG4gICAgICB9O1xuICAgICAgdGhpcy4kZWwuYWRkRXZlbnRMaXN0ZW5lcihrZXksIHRoaXMuX2hhbmRsZXJzW2V2ZW50VG9IYW5kbGVyKGtleSldKTtcbiAgICB9KTtcbiAgfSxcblxuICBiZWZvcmVEZXN0cm95KCkge1xuICAgIE9iamVjdC5rZXlzKHRoaXMuX2hhbmRsZXJzKS5mb3JFYWNoKGtleSA9PiB7XG4gICAgICB0aGlzLiRlbC5yZW1vdmVFdmVudExpc3RlbmVyKGtleSwgdGhpcy5faGFuZGxlcnNba2V5XSk7XG4gICAgfSk7XG4gICAgdGhpcy5faGFuZGxlcnMgPSBudWxsO1xuICB9XG59O1xuXG5leHBvcnQgeyBkZXJpdmVEQkIsIGRlcml2ZUV2ZW50cyB9O1xuIiwiLyogUHJpdmF0ZSAqL1xuY29uc3QgX3RvZ2dsZVZpc2liaWxpdHkgPSBmdW5jdGlvbigpIHtcbiAgaWYgKHR5cGVvZiB0aGlzLnZpc2libGUgPT09ICdib29sZWFuJyAmJiB0aGlzLnZpc2libGUgIT09IHRoaXMuJGVsLnZpc2libGUpIHtcbiAgICB0aGlzLiRlbFt0aGlzLnZpc2libGUgPyAnc2hvdycgOiAnaGlkZSddLmNhbGwodGhpcy4kZWwsIHRoaXMubm9ybWFsaXplZE9wdGlvbnMgfHwgdGhpcy5vcHRpb25zKTtcbiAgfVxufTtcbmNvbnN0IF90ZWxlcG9ydCA9IGZ1bmN0aW9uKCkge1xuICBpZiAoIXRoaXMuX2lzRGVzdHJveWVkICYmICghdGhpcy4kZWwucGFyZW50Tm9kZSB8fCB0aGlzLiRlbC5wYXJlbnROb2RlICE9PSBkb2N1bWVudC5ib2R5KSkge1xuICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQodGhpcy4kZWwpO1xuICB9XG59O1xuY29uc3QgX3VubW91bnQgPSBmdW5jdGlvbigpIHtcbiAgaWYgKHRoaXMuJGVsLnZpc2libGUgPT09IHRydWUpIHtcbiAgICB0aGlzLiRlbC5oaWRlKCkudGhlbigoKSA9PiB0aGlzLiRlbC5yZW1vdmUoKSk7XG4gIH0gZWxzZSB7XG4gICAgdGhpcy4kZWwucmVtb3ZlKCk7XG4gIH1cbn07XG5cbi8qIFB1YmxpYyAqL1xuLy8gQ29tcG9uZW50cyB0aGF0IGNhbiBiZSBzaG93biBvciBoaWRkZW5cbmNvbnN0IGhpZGFibGUgPSB7XG4gIHByb3BzOiB7XG4gICAgdmlzaWJsZToge1xuICAgICAgdHlwZTogQm9vbGVhbixcbiAgICAgIGRlZmF1bHQ6IHVuZGVmaW5lZCAvLyBBdm9pZCBjYXN0aW5nIHRvIGZhbHNlXG4gICAgfVxuICB9LFxuXG4gIHdhdGNoOiB7XG4gICAgdmlzaWJsZSgpIHtcbiAgICAgIF90b2dnbGVWaXNpYmlsaXR5LmNhbGwodGhpcyk7XG4gICAgfVxuICB9LFxuXG4gIG1vdW50ZWQoKSB7XG4gICAgdGhpcy4kbmV4dFRpY2soKCkgPT4gX3RvZ2dsZVZpc2liaWxpdHkuY2FsbCh0aGlzKSk7XG4gIH0sXG5cbiAgYWN0aXZhdGVkKCkge1xuICAgIHRoaXMuJG5leHRUaWNrKCgpID0+IF90b2dnbGVWaXNpYmlsaXR5LmNhbGwodGhpcykpO1xuICB9XG59O1xuXG4vLyBDb21wb25lbnRzIHdpdGggJ29wdGlvbnMnIHByb3BlcnR5XG5jb25zdCBoYXNPcHRpb25zID0ge1xuICBwcm9wczoge1xuICAgIG9wdGlvbnM6IHtcbiAgICAgIHR5cGU6IE9iamVjdCxcbiAgICAgIGRlZmF1bHQoKSB7XG4gICAgICAgIHJldHVybiB7fTtcbiAgICAgIH1cbiAgICB9XG4gIH1cbn07XG5cbi8vIFByb3ZpZGVzIGl0c2VsZiB0byBpdHMgZGVzY2VuZGFudHNcbmNvbnN0IHNlbGZQcm92aWRlciA9IHtcbiAgcHJvdmlkZSgpIHtcbiAgICByZXR1cm4ge1xuICAgICAgW3RoaXMuJG9wdGlvbnMuX2NvbXBvbmVudFRhZy5zbGljZSg2KV06IHRoaXNcbiAgICB9XG4gIH1cbn07XG5cbi8vIENvbW1vbiBldmVudCBmb3IgRGlhbG9nc1xuY29uc3QgZGlhbG9nQ2FuY2VsID0ge1xuICBtb3VudGVkKCkge1xuICAgIHRoaXMuJG9uKCdkaWFsb2ctY2FuY2VsJywgKCkgPT4gdGhpcy4kZW1pdCgndXBkYXRlOnZpc2libGUnLCBmYWxzZSkpO1xuICB9XG59O1xuXG4vLyBNb3ZlcyB0aGUgZWxlbWVudCB0byBhIGdsb2JhbCBwb3NpdGlvblxuY29uc3QgcG9ydGFsID0ge1xuICBtb3VudGVkKCkge1xuICAgIF90ZWxlcG9ydC5jYWxsKHRoaXMpO1xuICB9LFxuICB1cGRhdGVkKCkge1xuICAgIF90ZWxlcG9ydC5jYWxsKHRoaXMpO1xuICB9LFxuICBhY3RpdmF0ZWQoKSB7XG4gICAgX3RlbGVwb3J0LmNhbGwodGhpcyk7XG4gIH0sXG4gIGRlYWN0aXZhdGVkKCkge1xuICAgIF91bm1vdW50LmNhbGwodGhpcyk7XG4gIH0sXG4gIGJlZm9yZURlc3Ryb3koKSB7XG4gICAgX3VubW91bnQuY2FsbCh0aGlzKTtcbiAgfVxufTtcblxuY29uc3QgbW9kaWZpZXIgPSB7XG4gIHByb3BzOiB7XG4gICAgbW9kaWZpZXI6IHtcbiAgICAgIHR5cGU6IFtTdHJpbmcsIEFycmF5LCBPYmplY3RdXG4gICAgfSxcbiAgfSxcblxuICBjb21wdXRlZDoge1xuICAgIG5vcm1hbGl6ZWRNb2RpZmllcigpIHtcbiAgICAgIGNvbnN0IG1vZGlmaWVyID0gdGhpcy5tb2RpZmllcjtcblxuICAgICAgaWYgKHR5cGVvZiBtb2RpZmllciA9PT0gJ3N0cmluZycpIHtcbiAgICAgICByZXR1cm4gbW9kaWZpZXI7XG4gICAgICB9XG5cbiAgICAgIGlmIChBcnJheS5pc0FycmF5KG1vZGlmaWVyKSkge1xuICAgICAgICByZXR1cm4gbW9kaWZpZXIuam9pbignICcpO1xuICAgICAgfVxuXG4gICAgICBpZiAodHlwZW9mIG1vZGlmaWVyID09PSAnb2JqZWN0Jykge1xuICAgICAgICByZXR1cm4gT2JqZWN0LmtleXMobW9kaWZpZXIpXG4gICAgICAgICAgLnJlZHVjZSgoYWNjLCBrZXkpID0+IChhY2MgKyAobW9kaWZpZXJba2V5XSA/IGAgJHtrZXl9YCA6ICcnKSksICcnKVxuICAgICAgICAgIC50cmltKCk7XG4gICAgICB9XG5cbiAgICAgIHJldHVybiBmYWxzZTtcbiAgICB9XG4gIH1cbn07XG5cbmV4cG9ydCB7IGhpZGFibGUsIGhhc09wdGlvbnMsIHNlbGZQcm92aWRlciwgZGlhbG9nQ2FuY2VsLCBwb3J0YWwsIG1vZGlmaWVyIH07XG4iLCIvKiBQcml2YXRlICovXG5jb25zdCBtb2RlbCA9IHtcbiAgcHJvcDogJ21vZGVsUHJvcCcsXG4gIGV2ZW50OiAnbW9kZWxFdmVudCdcbn07XG5cbi8qIFB1YmxpYyAqL1xuXG4vLyBHZW5lcmljIGlucHV0XG5jb25zdCBtb2RlbElucHV0ID0ge1xuICBtb2RlbCxcbiAgcHJvcHM6IHtcbiAgICBbbW9kZWwucHJvcF06IFtOdW1iZXIsIFN0cmluZ10sXG4gICAgW21vZGVsLmV2ZW50XToge1xuICAgICAgdHlwZTogU3RyaW5nLFxuICAgICAgZGVmYXVsdDogJ2lucHV0J1xuICAgIH1cbiAgfSxcblxuICBtZXRob2RzOiB7XG4gICAgX3VwZGF0ZVZhbHVlKCkge1xuICAgICAgaWYgKHRoaXNbbW9kZWwucHJvcF0gIT09IHVuZGVmaW5lZCAmJiB0aGlzLiRlbC52YWx1ZSAhPT0gdGhpc1ttb2RlbC5wcm9wXSkge1xuICAgICAgICB0aGlzLiRlbC52YWx1ZSA9IHRoaXNbbW9kZWwucHJvcF07XG4gICAgICB9XG4gICAgfSxcbiAgICBfb25Nb2RlbEV2ZW50KGV2ZW50KSB7XG4gICAgICB0aGlzLiRlbWl0KG1vZGVsLmV2ZW50LCBldmVudC50YXJnZXQudmFsdWUpO1xuICAgIH1cbiAgfSxcblxuICB3YXRjaDoge1xuICAgIFttb2RlbC5wcm9wXSgpIHtcbiAgICAgIHRoaXMuX3VwZGF0ZVZhbHVlKCk7XG4gICAgfVxuICB9LFxuXG4gIG1vdW50ZWQoKSB7XG4gICAgdGhpcy5fdXBkYXRlVmFsdWUoKTtcbiAgICB0aGlzLiRlbC5hZGRFdmVudExpc3RlbmVyKHRoaXNbbW9kZWwuZXZlbnRdLCB0aGlzLl9vbk1vZGVsRXZlbnQpO1xuICB9LFxuICBiZWZvcmVEZXN0cm95KCkge1xuICAgIHRoaXMuJGVsLnJlbW92ZUV2ZW50TGlzdGVuZXIodGhpc1ttb2RlbC5ldmVudF0sIHRoaXMuX29uTW9kZWxFdmVudCk7XG4gIH1cbn07XG5cbi8vIENoZWNrYWJsZSBpbnB1dHNcbmNvbnN0IG1vZGVsQ2hlY2tib3ggPSB7XG4gIG1peGluczogW21vZGVsSW5wdXRdLFxuXG4gIHByb3BzOiB7XG4gICAgW21vZGVsLnByb3BdOiBbQXJyYXksIEJvb2xlYW5dLFxuICAgIFttb2RlbC5ldmVudF06IHtcbiAgICAgIHR5cGU6IFN0cmluZyxcbiAgICAgIGRlZmF1bHQ6ICdjaGFuZ2UnXG4gICAgfVxuICB9LFxuXG4gIG1ldGhvZHM6IHtcbiAgICBfdXBkYXRlVmFsdWUoKSB7XG4gICAgICBpZiAodGhpc1ttb2RlbC5wcm9wXSBpbnN0YW5jZW9mIEFycmF5KSB7XG4gICAgICAgIHRoaXMuJGVsLmNoZWNrZWQgPSB0aGlzW21vZGVsLnByb3BdLmluZGV4T2YodGhpcy4kZWwudmFsdWUpID49IDA7XG4gICAgICB9IGVsc2Uge1xuICAgICAgICB0aGlzLiRlbC5jaGVja2VkID0gdGhpc1ttb2RlbC5wcm9wXTtcbiAgICAgIH1cbiAgICB9LFxuICAgIF9vbk1vZGVsRXZlbnQoZXZlbnQpIHtcbiAgICAgIGNvbnN0IHsgdmFsdWUsIGNoZWNrZWQgfSA9IGV2ZW50LnRhcmdldDtcbiAgICAgIGxldCBuZXdWYWx1ZTtcblxuICAgICAgaWYgKHRoaXNbbW9kZWwucHJvcF0gaW5zdGFuY2VvZiBBcnJheSkge1xuICAgICAgICAvLyBJcyBBcnJheVxuICAgICAgICBjb25zdCBpbmRleCA9IHRoaXNbbW9kZWwucHJvcF0uaW5kZXhPZih2YWx1ZSk7XG4gICAgICAgIGNvbnN0IGluY2x1ZGVkID0gaW5kZXggPj0gMDtcblxuICAgICAgICBpZiAoaW5jbHVkZWQgJiYgIWNoZWNrZWQpIHtcbiAgICAgICAgICBuZXdWYWx1ZSA9IFtcbiAgICAgICAgICAgIC4uLnRoaXNbbW9kZWwucHJvcF0uc2xpY2UoMCwgaW5kZXgpLFxuICAgICAgICAgICAgLi4udGhpc1ttb2RlbC5wcm9wXS5zbGljZShpbmRleCArIDEsIHRoaXNbbW9kZWwucHJvcF0ubGVuZ3RoKVxuICAgICAgICAgIF07XG4gICAgICAgIH1cblxuICAgICAgICBpZiAoIWluY2x1ZGVkICYmIGNoZWNrZWQpIHtcbiAgICAgICAgICBuZXdWYWx1ZSA9IFsgLi4udGhpc1ttb2RlbC5wcm9wXSwgdmFsdWUgXTtcbiAgICAgICAgfVxuXG4gICAgICB9IGVsc2Uge1xuICAgICAgICAvLyBJcyBCb29sZWFuXG4gICAgICAgIG5ld1ZhbHVlID0gY2hlY2tlZDtcbiAgICAgIH1cblxuICAgICAgLy8gRW1pdCBpZiB2YWx1ZSBjaGFuZ2VkXG4gICAgICBuZXdWYWx1ZSAhPT0gdW5kZWZpbmVkICYmIHRoaXMuJGVtaXQobW9kZWwuZXZlbnQsIG5ld1ZhbHVlKTtcbiAgICB9XG4gIH1cbn07XG5cbi8vIFJhZGlvIGlucHV0XG5jb25zdCBtb2RlbFJhZGlvID0ge1xuICBtaXhpbnM6IFttb2RlbElucHV0XSxcbiAgcHJvcHM6IHtcbiAgICBbbW9kZWwuZXZlbnRdOiB7XG4gICAgICB0eXBlOiBTdHJpbmcsXG4gICAgICBkZWZhdWx0OiAnY2hhbmdlJ1xuICAgIH1cbiAgfSxcblxuICBtZXRob2RzOiB7XG4gICAgX3VwZGF0ZVZhbHVlKCkge1xuICAgICAgdGhpcy4kZWwuY2hlY2tlZCA9IHRoaXNbbW9kZWwucHJvcF0gPT09IHRoaXMuJGVsLnZhbHVlO1xuICAgIH0sXG4gICAgX29uTW9kZWxFdmVudChldmVudCkge1xuICAgICAgY29uc3QgeyB2YWx1ZSwgY2hlY2tlZCB9ID0gZXZlbnQudGFyZ2V0O1xuICAgICAgY2hlY2tlZCAmJiB0aGlzLiRlbWl0KG1vZGVsLmV2ZW50LCB2YWx1ZSk7XG4gICAgfVxuICB9XG59O1xuXG5leHBvcnQgeyBtb2RlbElucHV0LCBtb2RlbENoZWNrYm94LCBtb2RlbFJhZGlvIH07XG5cbiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy10b29sYmFyIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXRvb2xiYXI+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy10b29sYmFyJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy10b29sYmFyJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWJvdHRvbS10b29sYmFyIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWJvdHRvbS10b29sYmFyPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtYm90dG9tLXRvb2xiYXInO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWJvdHRvbS10b29sYmFyJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLXRvb2xiYXItYnV0dG9uIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXRvb2xiYXItYnV0dG9uPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtdG9vbGJhci1idXR0b24nO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLXRvb2xiYXItYnV0dG9uJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWFsZXJ0LWRpYWxvZy1idXR0b24gdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtYWxlcnQtZGlhbG9nLWJ1dHRvbj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWFsZXJ0LWRpYWxvZy1idXR0b24nO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWFsZXJ0LWRpYWxvZy1idXR0b24nLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtYnV0dG9uIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWJ1dHRvbj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWJ1dHRvbic7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtYnV0dG9uJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWljb24gdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtaWNvbj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWljb24nO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWljb24nLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtY2FyZCB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1jYXJkPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtY2FyZCc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtY2FyZCcsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1saXN0IHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWxpc3Q+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1saXN0JztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1saXN0JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWxpc3QtaXRlbSB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1saXN0LWl0ZW0+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1saXN0LWl0ZW0nO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWxpc3QtaXRlbScsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RpZmllcl0sXG4gICAgcHJvcHM6IHtcbiAgICAgIGV4cGFuZGVkOiB7XG4gICAgICAgIHR5cGU6IEJvb2xlYW5cbiAgICAgIH1cbiAgICB9LFxuXG4gICAgd2F0Y2g6IHtcbiAgICAgIGV4cGFuZGVkKCkge1xuICAgICAgICBjb25zdCBhY3Rpb24gPSB0aGlzLmV4cGFuZGVkID8gJ3Nob3cnIDogJ2hpZGUnO1xuICAgICAgICB0aGlzLiRlbFthY3Rpb24gKyAnRXhwYW5zaW9uJ10oKTtcbiAgICAgIH1cbiAgICB9XG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtbGlzdC10aXRsZSB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1saXN0LXRpdGxlPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtbGlzdC10aXRsZSc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtbGlzdC10aXRsZScsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1saXN0LWhlYWRlciB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1saXN0LWhlYWRlcj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWxpc3QtaGVhZGVyJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1saXN0LWhlYWRlcicsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1yaXBwbGUgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtcmlwcGxlPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtcmlwcGxlJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLXJpcHBsZScsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLXJvdyB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1yb3c+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1yb3cnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtcm93JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHNdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtY29sIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWNvbD5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWNvbCc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1jb2wnLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50c11cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1wcm9ncmVzcy1iYXIgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtcHJvZ3Jlc3MtYmFyPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtcHJvZ3Jlc3MtYmFyJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1wcm9ncmVzcy1iYXInLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtcHJvZ3Jlc3MtY2lyY3VsYXIgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtcHJvZ3Jlc3MtY2lyY3VsYXI+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1wcm9ncmVzcy1jaXJjdWxhcic7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtcHJvZ3Jlc3MtY2lyY3VsYXInLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtY2Fyb3VzZWwtaXRlbSB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1jYXJvdXNlbC1pdGVtPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtY2Fyb3VzZWwtaXRlbSc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1jYXJvdXNlbC1pdGVtJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHNdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc3BsaXR0ZXItbWFzayB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1zcGxpdHRlci1tYXNrPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtc3BsaXR0ZXItbWFzayc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1zcGxpdHRlci1tYXNrJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHNdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc3BsaXR0ZXItY29udGVudCB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1zcGxpdHRlci1jb250ZW50PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtc3BsaXR0ZXItY29udGVudCc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1zcGxpdHRlci1jb250ZW50JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHNdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc3BsaXR0ZXIgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtc3BsaXR0ZXI+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1zcGxpdHRlcic7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgc2VsZlByb3ZpZGVyLCBkZXJpdmVEQkIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtc3BsaXR0ZXInLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgc2VsZlByb3ZpZGVyLCBkZXJpdmVEQkJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc3dpdGNoIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXN3aXRjaD5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLXN3aXRjaCc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kZWxDaGVja2JveCwgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtc3dpdGNoJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGVsQ2hlY2tib3gsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWNoZWNrYm94IHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWNoZWNrYm94PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtY2hlY2tib3gnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGVsQ2hlY2tib3gsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWNoZWNrYm94JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGVsQ2hlY2tib3gsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWlucHV0IHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLWlucHV0PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtaW5wdXQnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGVsSW5wdXQsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWlucHV0JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGVsSW5wdXQsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLXNlYXJjaC1pbnB1dCB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1zZWFyY2gtaW5wdXQ+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1zZWFyY2gtaW5wdXQnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIG1vZGVsSW5wdXQsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLXNlYXJjaC1pbnB1dCcsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RlbElucHV0LCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1yYW5nZSB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1yYW5nZT5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLXJhbmdlJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RlbElucHV0LCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1yYW5nZScsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RlbElucHV0LCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1yYWRpbyB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1yYWRpbz5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLXJhZGlvJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBtb2RlbFJhZGlvLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1yYWRpbycsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RlbFJhZGlvLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1mYWIgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtZmFiPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtZmFiJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBoaWRhYmxlLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1mYWInLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgaGlkYWJsZSwgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc3BlZWQtZGlhbC1pdGVtIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXNwZWVkLWRpYWwtaXRlbT5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLXNwZWVkLWRpYWwtaXRlbSc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtc3BlZWQtZGlhbC1pdGVtJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLWRpYWxvZyB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1kaWFsb2c+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICAvKiBUaGlzIGZpbGUgd2FzIGdlbmVyYXRlZCBhdXRvbWF0aWNhbGx5IGJ5ICdnZW5lcmF0ZS1jb21wb25lbnRzJyB0YXNrIGluIGJpbmRpbmdzL3Z1ZS9ndWxwZmlsZS5iYWJlbC5qcyAqL1xuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1kaWFsb2cnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIGhpZGFibGUsIGhhc09wdGlvbnMsIGRpYWxvZ0NhbmNlbCwgZGVyaXZlREJCLCBwb3J0YWwsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWRpYWxvZycsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBoaWRhYmxlLCBoYXNPcHRpb25zLCBkaWFsb2dDYW5jZWwsIGRlcml2ZURCQiwgcG9ydGFsLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1hY3Rpb24tc2hlZXQgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxzbG90Pjwvc2xvdD5cbiAgPC9vbnMtYWN0aW9uLXNoZWV0PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtYWN0aW9uLXNoZWV0JztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBoaWRhYmxlLCBoYXNPcHRpb25zLCBkaWFsb2dDYW5jZWwsIGRlcml2ZURCQiwgcG9ydGFsLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1hY3Rpb24tc2hlZXQnLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgaGlkYWJsZSwgaGFzT3B0aW9ucywgZGlhbG9nQ2FuY2VsLCBkZXJpdmVEQkIsIHBvcnRhbCwgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtYWN0aW9uLXNoZWV0LWJ1dHRvbiB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1hY3Rpb24tc2hlZXQtYnV0dG9uPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtYWN0aW9uLXNoZWV0LWJ1dHRvbic7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtYWN0aW9uLXNoZWV0LWJ1dHRvbicsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1tb2RhbCB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1tb2RhbD5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIC8qIFRoaXMgZmlsZSB3YXMgZ2VuZXJhdGVkIGF1dG9tYXRpY2FsbHkgYnkgJ2dlbmVyYXRlLWNvbXBvbmVudHMnIHRhc2sgaW4gYmluZGluZ3MvdnVlL2d1bHBmaWxlLmJhYmVsLmpzICovXG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLW1vZGFsJztcbiAgaW1wb3J0IHsgZGVyaXZlRXZlbnRzLCBoaWRhYmxlLCBoYXNPcHRpb25zLCBkZXJpdmVEQkIsIHBvcnRhbCwgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtbW9kYWwnLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgaGlkYWJsZSwgaGFzT3B0aW9ucywgZGVyaXZlREJCLCBwb3J0YWwsIG1vZGlmaWVyXVxuICB9O1xuPC9zY3JpcHQ+IiwiPHRlbXBsYXRlPlxuICA8b25zLXRvYXN0IHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIiA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXRvYXN0PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgLyogVGhpcyBmaWxlIHdhcyBnZW5lcmF0ZWQgYXV0b21hdGljYWxseSBieSAnZ2VuZXJhdGUtY29tcG9uZW50cycgdGFzayBpbiBiaW5kaW5ncy92dWUvZ3VscGZpbGUuYmFiZWwuanMgKi9cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtdG9hc3QnO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIGhpZGFibGUsIGhhc09wdGlvbnMsIGRlcml2ZURCQiwgcG9ydGFsLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy10b2FzdCcsXG4gICAgbWl4aW5zOiBbZGVyaXZlRXZlbnRzLCBoaWRhYmxlLCBoYXNPcHRpb25zLCBkZXJpdmVEQkIsIHBvcnRhbCwgbW9kaWZpZXJdXG4gIH07XG48L3NjcmlwdD4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtcG9wb3ZlciB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1wb3BvdmVyPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtcG9wb3Zlcic7XG4gIGltcG9ydCB7IGhpZGFibGUsIGhhc09wdGlvbnMsIGRpYWxvZ0NhbmNlbCwgZGVyaXZlRXZlbnRzLCBkZXJpdmVEQkIsIHBvcnRhbCB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1wb3BvdmVyJyxcbiAgICBtaXhpbnM6IFtoaWRhYmxlLCBoYXNPcHRpb25zLCBkaWFsb2dDYW5jZWwsIGRlcml2ZUV2ZW50cywgZGVyaXZlREJCLCBwb3J0YWxdLFxuXG4gICAgcHJvcHM6IHtcbiAgICAgIHRhcmdldDoge1xuICAgICAgICB2YWxpZGF0b3IodmFsdWUpIHtcbiAgICAgICAgICByZXR1cm4gdmFsdWUuX2lzVnVlIHx8IHR5cGVvZiB2YWx1ZSA9PT0gJ3N0cmluZycgfHwgdmFsdWUgaW5zdGFuY2VvZiBFdmVudCB8fCB2YWx1ZSBpbnN0YW5jZW9mIEhUTUxFbGVtZW50O1xuICAgICAgICB9XG4gICAgICB9XG4gICAgfSxcblxuICAgIGNvbXB1dGVkOiB7XG4gICAgICBub3JtYWxpemVkVGFyZ2V0KCkge1xuICAgICAgICBpZiAodGhpcy50YXJnZXQgJiYgdGhpcy50YXJnZXQuX2lzVnVlKSB7XG4gICAgICAgICAgcmV0dXJuIHRoaXMudGFyZ2V0LiRlbDtcbiAgICAgICAgfVxuICAgICAgICByZXR1cm4gdGhpcy50YXJnZXQ7XG4gICAgICB9LFxuICAgICAgbm9ybWFsaXplZE9wdGlvbnMoKSB7XG4gICAgICAgIGlmICh0aGlzLnRhcmdldCkge1xuICAgICAgICAgIHJldHVybiB7XG4gICAgICAgICAgICB0YXJnZXQ6IHRoaXMubm9ybWFsaXplZFRhcmdldCxcbiAgICAgICAgICAgIC4uLnRoaXMub3B0aW9uc1xuICAgICAgICAgIH07XG4gICAgICAgIH1cbiAgICAgICAgcmV0dXJuIHRoaXMub3B0aW9ucztcbiAgICAgIH1cbiAgICB9XG4gIH07XG48L3NjcmlwdD5cbiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1hbGVydC1kaWFsb2cgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiPlxuICAgIDxkaXYgY2xhc3M9XCJhbGVydC1kaWFsb2ctdGl0bGVcIj5cbiAgICAgIDxzbG90IG5hbWU9XCJ0aXRsZVwiPnt7dGl0bGV9fTwvc2xvdD5cbiAgICA8L2Rpdj5cbiAgICA8ZGl2IGNsYXNzPVwiYWxlcnQtZGlhbG9nLWNvbnRlbnRcIj5cbiAgICAgIDxzbG90Pjwvc2xvdD5cbiAgICA8L2Rpdj5cbiAgICA8ZGl2IGNsYXNzPVwiYWxlcnQtZGlhbG9nLWZvb3RlclwiPlxuICAgICAgPHNsb3QgbmFtZT1cImZvb3RlclwiPlxuICAgICAgICA8b25zLWFsZXJ0LWRpYWxvZy1idXR0b24gdi1mb3I9XCIoaGFuZGxlciwga2V5KSBpbiBmb290ZXJcIiA6a2V5PVwia2V5XCIgQGNsaWNrPVwiaGFuZGxlclwiPnt7a2V5fX08L29ucy1hbGVydC1kaWFsb2ctYnV0dG9uPlxuICAgICAgPC9zbG90PlxuICAgIDwvZGl2PlxuICA8L29ucy1hbGVydC1kaWFsb2c+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1hbGVydC1kaWFsb2cnO1xuICBpbXBvcnQgeyBoaWRhYmxlLCBoYXNPcHRpb25zLCBkaWFsb2dDYW5jZWwsIGRlcml2ZUV2ZW50cywgZGVyaXZlREJCLCBwb3J0YWwsIG1vZGlmaWVyIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLWFsZXJ0LWRpYWxvZycsXG4gICAgbWl4aW5zOiBbaGlkYWJsZSwgaGFzT3B0aW9ucywgZGlhbG9nQ2FuY2VsLCBkZXJpdmVFdmVudHMsIGRlcml2ZURCQiwgcG9ydGFsLCBtb2RpZmllcl0sXG5cbiAgICBwcm9wczoge1xuICAgICAgdGl0bGU6IHtcbiAgICAgICAgdHlwZTogU3RyaW5nXG4gICAgICB9LFxuICAgICAgZm9vdGVyOiB7XG4gICAgICAgIHR5cGU6IE9iamVjdCxcbiAgICAgICAgdmFsaWRhdG9yKHZhbHVlKSB7XG4gICAgICAgICAgcmV0dXJuIE9iamVjdC5rZXlzKHZhbHVlKS5ldmVyeShrZXkgPT4gdmFsdWVba2V5XSBpbnN0YW5jZW9mIEZ1bmN0aW9uKTtcbiAgICAgICAgfVxuICAgICAgfVxuICAgIH1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8b25zLXNwZWVkLWRpYWwgOm9uLWNsaWNrLnByb3A9XCJhY3Rpb25cIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXNwZWVkLWRpYWw+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1zcGVlZC1kaWFsJztcbiAgaW1wb3J0IHsgaGlkYWJsZSwgZGVyaXZlRXZlbnRzIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLXNwZWVkLWRpYWwnLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50cywgaGlkYWJsZV0sXG5cbiAgICBwcm9wczoge1xuICAgICAgb3Blbjoge1xuICAgICAgICB0eXBlOiBCb29sZWFuLFxuICAgICAgICBkZWZhdWx0OiB1bmRlZmluZWRcbiAgICAgIH1cbiAgICB9LFxuXG4gICAgbWV0aG9kczoge1xuICAgICAgYWN0aW9uKCkge1xuICAgICAgICBsZXQgcnVuRGVmYXVsdCA9IHRydWU7XG4gICAgICAgIHRoaXMuJGVtaXQoJ2NsaWNrJywgeyBwcmV2ZW50RGVmYXVsdDogKCkgPT4gcnVuRGVmYXVsdCA9IGZhbHNlIH0pO1xuXG4gICAgICAgIGlmIChydW5EZWZhdWx0KSB7XG4gICAgICAgICAgdGhpcy4kZWwudG9nZ2xlSXRlbXMoKTtcbiAgICAgICAgfVxuICAgICAgfSxcbiAgICAgIF9zaG91bGRVcGRhdGUoKSB7XG4gICAgICAgIHJldHVybiB0aGlzLm9wZW4gIT09IHVuZGVmaW5lZCAmJiB0aGlzLm9wZW4gIT09IHRoaXMuJGVsLmlzT3BlbigpO1xuICAgICAgfSxcbiAgICAgIF91cGRhdGVUb2dnbGUoKSB7XG4gICAgICAgIHRoaXMuX3Nob3VsZFVwZGF0ZSgpICYmIHRoaXMuJGVsW3RoaXMub3BlbiA/ICdzaG93SXRlbXMnIDogJ2hpZGVJdGVtcyddLmNhbGwodGhpcy4kZWwpO1xuICAgICAgfVxuICAgIH0sXG5cbiAgICB3YXRjaDoge1xuICAgICAgb3BlbigpIHtcbiAgICAgICAgdGhpcy5fdXBkYXRlVG9nZ2xlKCk7XG4gICAgICB9XG4gICAgfSxcblxuICAgIG1vdW50ZWQoKSB7XG4gICAgICB0aGlzLiRvbihbJ29wZW4nLCAnY2xvc2UnXSwgKCkgPT4gdGhpcy5fc2hvdWxkVXBkYXRlKCkgJiYgdGhpcy4kZW1pdCgndXBkYXRlOm9wZW4nLCB0aGlzLiRlbC5pc09wZW4oKSkpO1xuXG4gICAgICB0aGlzLl91cGRhdGVUb2dnbGUoKTtcbiAgICB9XG4gIH07XG48L3NjcmlwdD5cbiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1jYXJvdXNlbFxuICAgIDpvbi1zd2lwZS5wcm9wPVwib25Td2lwZVwiXG4gICAgOmluaXRpYWwtaW5kZXg9XCJpbmRleFwiXG4gICAgQHBvc3RjaGFuZ2Uuc2VsZj1cIiRlbWl0KCd1cGRhdGU6aW5kZXgnLCAkZXZlbnQuYWN0aXZlSW5kZXgpXCJcbiAgICB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCJcbiAgPlxuICAgIDxkaXY+XG4gICAgICA8c2xvdD48L3Nsb3Q+XG4gICAgPC9kaXY+XG4gICAgPGRpdj48L2Rpdj5cbiAgPC9vbnMtY2Fyb3VzZWw+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1jYXJvdXNlbCc7XG4gIGltcG9ydCB7IGhhc09wdGlvbnMsIGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1jYXJvdXNlbCcsXG4gICAgbWl4aW5zOiBbaGFzT3B0aW9ucywgZGVyaXZlRXZlbnRzXSxcblxuICAgIHByb3BzOiB7XG4gICAgICBpbmRleDoge1xuICAgICAgICB0eXBlOiBOdW1iZXJcbiAgICAgIH0sXG4gICAgICBvblN3aXBlOiB7XG4gICAgICAgIHR5cGU6IEZ1bmN0aW9uXG4gICAgICB9XG4gICAgfSxcblxuICAgIHdhdGNoOiB7XG4gICAgICBpbmRleCgpIHtcbiAgICAgICAgaWYgKHRoaXMuaW5kZXggIT09IHRoaXMuJGVsLmdldEFjdGl2ZUluZGV4KCkpIHtcbiAgICAgICAgICB0aGlzLiRlbC5zZXRBY3RpdmVJbmRleCh0aGlzLmluZGV4LCB0aGlzLm9wdGlvbnMpO1xuICAgICAgICB9XG4gICAgICB9XG4gICAgfVxuICB9O1xuPC9zY3JpcHQ+XG4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtdGFiIDphY3RpdmU9XCJhY3RpdmVcIiA6b24tY2xpY2sucHJvcD1cImFjdGlvblwiPlxuICA8L29ucy10YWI+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy10YWInO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtdGFiJyxcbiAgICBpbmplY3Q6IFsndGFiYmFyJ10sXG5cbiAgICBwcm9wczoge1xuICAgICAgcGFnZTogeyB9LFxuICAgICAgcHJvcHM6IHsgfSxcbiAgICAgIGFjdGl2ZToge1xuICAgICAgICB0eXBlOiBCb29sZWFuXG4gICAgICB9XG4gICAgfSxcblxuICAgIG1ldGhvZHM6IHtcbiAgICAgIGFjdGlvbigpIHtcbiAgICAgICAgbGV0IHJ1bkRlZmF1bHQgPSB0cnVlO1xuICAgICAgICB0aGlzLiRlbWl0KCdjbGljaycsIHsgcHJldmVudERlZmF1bHQ6ICgpID0+IHJ1bkRlZmF1bHQgPSBmYWxzZSB9KTtcblxuICAgICAgICBpZiAocnVuRGVmYXVsdCkge1xuICAgICAgICAgIHRoaXMudGFiYmFyLiRlbC5zZXRBY3RpdmVUYWIodGhpcy4kZWwuaW5kZXgsIHsgcmVqZWN0OiBmYWxzZSwgLi4udGhpcy50YWJiYXIub3B0aW9ucyB9KTtcbiAgICAgICAgfVxuICAgICAgfVxuICAgIH0sXG5cbiAgICB3YXRjaDoge1xuICAgICAgYWN0aXZlKCkge1xuICAgICAgICB0aGlzLiRlbC5zZXRBY3RpdmUodGhpcy5hY3RpdmUpO1xuICAgICAgfVxuICAgIH1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8b25zLXRhYmJhclxuICAgIDpvbi1zd2lwZS5wcm9wPVwib25Td2lwZVwiXG4gICAgOmFjdGl2ZUluZGV4PVwiaW5kZXhcIlxuICAgIDptb2RpZmllcj1cIm5vcm1hbGl6ZWRNb2RpZmllclwiXG4gICAgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiXG4gICAgQHByZWNoYW5nZS5zZWxmPVwiJG5leHRUaWNrKCgpID0+ICEkZXZlbnQuZGV0YWlsLmNhbmNlbGVkICYmICRlbWl0KCd1cGRhdGU6aW5kZXgnLCAkZXZlbnQuaW5kZXgpKVwiXG4gID5cbiAgICA8ZGl2IGNsYXNzPVwidGFiYmFyX19jb250ZW50XCI+XG4gICAgICA8ZGl2PlxuICAgICAgICA8c2xvdCBuYW1lPVwicGFnZXNcIj5cbiAgICAgICAgICA8Y29tcG9uZW50IHYtZm9yPVwidGFiIGluIHRhYnNcIiB2LWJpbmQ9XCJ0YWIucHJvcHNcIiA6aXM9XCJ0YWIucGFnZVwiIDprZXk9XCIodGFiLnBhZ2Uua2V5IHx8IHRhYi5wYWdlLm5hbWUgfHwgX3RhYktleSh0YWIpKVwiIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIj48L2NvbXBvbmVudD5cbiAgICAgICAgPC9zbG90PlxuICAgICAgPC9kaXY+XG4gICAgICA8ZGl2PjwvZGl2PlxuICAgIDwvZGl2PlxuICAgIDxkaXYgY2xhc3M9XCJ0YWJiYXJcIiA6c3R5bGU9XCJ0YWJiYXJTdHlsZVwiPlxuICAgICAgPHNsb3Q+XG4gICAgICAgIDx2LW9ucy10YWIgdi1mb3I9XCJ0YWIgaW4gdGFic1wiIHYtYmluZD1cInRhYlwiIDprZXk9XCJfdGFiS2V5KHRhYilcIj48L3Ytb25zLXRhYj5cbiAgICAgIDwvc2xvdD5cbiAgICAgIDxkaXYgY2xhc3M9XCJ0YWJiYXJfX2JvcmRlclwiPjwvZGl2PlxuICAgIDwvZGl2PlxuICA8L29ucy10YWJiYXI+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy10YWJiYXInO1xuICBpbXBvcnQgeyBkZXJpdmVFdmVudHMsIGhhc09wdGlvbnMsIGhpZGFibGUsIHNlbGZQcm92aWRlciwgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtdGFiYmFyJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIGhhc09wdGlvbnMsIGhpZGFibGUsIHNlbGZQcm92aWRlciwgbW9kaWZpZXJdLFxuXG4gICAgcHJvcHM6IHtcbiAgICAgIGluZGV4OiB7XG4gICAgICAgIHR5cGU6IE51bWJlclxuICAgICAgfSxcbiAgICAgIHRhYnM6IHtcbiAgICAgICAgdHlwZTogQXJyYXksXG4gICAgICAgIHZhbGlkYXRvcih2YWx1ZSkge1xuICAgICAgICAgIHJldHVybiB2YWx1ZS5ldmVyeSh0YWIgPT4gWydpY29uJywgJ2xhYmVsJywgJ3BhZ2UnXS5zb21lKHByb3AgPT4gISFPYmplY3QuZ2V0T3duUHJvcGVydHlEZXNjcmlwdG9yKHRhYiwgcHJvcCkpKTtcbiAgICAgICAgfVxuICAgICAgfSxcbiAgICAgIG9uU3dpcGU6IHtcbiAgICAgICAgdHlwZTogRnVuY3Rpb25cbiAgICAgIH0sXG4gICAgICB0YWJiYXJTdHlsZToge1xuICAgICAgICB0eXBlOiBudWxsXG4gICAgICB9XG4gICAgfSxcblxuICAgIG1ldGhvZHM6IHtcbiAgICAgIF90YWJLZXkodGFiKSB7XG4gICAgICAgIHJldHVybiB0YWIua2V5IHx8IHRhYi5sYWJlbCB8fCB0YWIuaWNvbjtcbiAgICAgIH1cbiAgICB9LFxuXG4gICAgd2F0Y2g6IHtcbiAgICAgIGluZGV4KCkge1xuICAgICAgICBpZiAodGhpcy5pbmRleCAhPT0gdGhpcy4kZWwuZ2V0QWN0aXZlVGFiSW5kZXgoKSkge1xuICAgICAgICAgIHRoaXMuJGVsLnNldEFjdGl2ZVRhYih0aGlzLmluZGV4LCB7IHJlamVjdDogZmFsc2UsIC4uLnRoaXMub3B0aW9ucyB9KTtcbiAgICAgICAgfVxuICAgICAgfVxuICAgIH1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8b25zLWJhY2stYnV0dG9uIDpvbi1jbGljay5wcm9wPVwiYWN0aW9uXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1iYWNrLWJ1dHRvbj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLWJhY2stYnV0dG9uJztcbiAgaW1wb3J0IHsgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtYmFjay1idXR0b24nLFxuICAgIGluamVjdDogWyduYXZpZ2F0b3InXSxcbiAgICBtaXhpbnM6IFttb2RpZmllcl0sXG5cbiAgICBtZXRob2RzOiB7XG4gICAgICBhY3Rpb24oKSB7XG4gICAgICAgIGxldCBydW5EZWZhdWx0ID0gdHJ1ZTtcbiAgICAgICAgdGhpcy4kZW1pdCgnY2xpY2snLCB7IHByZXZlbnREZWZhdWx0OiAoKSA9PiBydW5EZWZhdWx0ID0gZmFsc2UgfSk7XG5cbiAgICAgICAgaWYgKHJ1bkRlZmF1bHQgJiYgdGhpcy5uYXZpZ2F0b3IucGFnZVN0YWNrLmxlbmd0aCA+IDEpIHtcbiAgICAgICAgICB0aGlzLm5hdmlnYXRvci5wb3BQYWdlKCk7XG4gICAgICAgIH1cbiAgICAgIH1cbiAgICB9XG4gIH07XG48L3NjcmlwdD5cbiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1uYXZpZ2F0b3IgQHBvc3Rwb3Auc2VsZj1cIl9jaGVja1N3aXBlXCIgOm9wdGlvbnMucHJvcD1cIm9wdGlvbnNcIiB2LW9uPVwidW5yZWNvZ25pemVkTGlzdGVuZXJzXCI+XG4gICAgPHNsb3Q+XG4gICAgICA8Y29tcG9uZW50XG4gICAgICAgIHYtZm9yPVwicGFnZSBpbiBwYWdlU3RhY2tcIlxuICAgICAgICA6aXM9XCJwYWdlXCJcbiAgICAgICAgOmtleT1cInBhZ2Uua2V5IHx8IHBhZ2UubmFtZVwiXG4gICAgICAgIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIlxuICAgICAgICB2LWJpbmQ9XCJwYWdlLm9uc05hdmlnYXRvclByb3BzXCJcbiAgICAgID48L2NvbXBvbmVudD5cbiAgICA8L3Nsb3Q+XG4gIDwvb25zLW5hdmlnYXRvcj5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLW5hdmlnYXRvcic7XG4gIGltcG9ydCB7IGhhc09wdGlvbnMsIHNlbGZQcm92aWRlciwgZGVyaXZlRXZlbnRzLCBkZXJpdmVEQkIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtbmF2aWdhdG9yJyxcbiAgICBtaXhpbnM6IFtoYXNPcHRpb25zLCBzZWxmUHJvdmlkZXIsIGRlcml2ZUV2ZW50cywgZGVyaXZlREJCXSxcblxuICAgIHByb3BzOiB7XG4gICAgICBwYWdlU3RhY2s6IHtcbiAgICAgICAgdHlwZTogQXJyYXksXG4gICAgICAgIHJlcXVpcmVkOiB0cnVlXG4gICAgICB9LFxuICAgICAgcG9wUGFnZToge1xuICAgICAgICB0eXBlOiBGdW5jdGlvbixcbiAgICAgICAgZGVmYXVsdCgpIHtcbiAgICAgICAgICB0aGlzLnBhZ2VTdGFjay5wb3AoKTtcbiAgICAgICAgfVxuICAgICAgfVxuICAgIH0sXG5cbiAgICBtZXRob2RzOiB7XG4gICAgICBpc1JlYWR5KCkge1xuICAgICAgICBpZiAodGhpcy5oYXNPd25Qcm9wZXJ0eSgnX3JlYWR5JykgJiYgdGhpcy5fcmVhZHkgaW5zdGFuY2VvZiBQcm9taXNlKSB7XG4gICAgICAgICAgcmV0dXJuIHRoaXMuX3JlYWR5O1xuICAgICAgICB9XG4gICAgICAgIHJldHVybiBQcm9taXNlLnJlc29sdmUoKTtcbiAgICAgIH0sXG4gICAgICBvbkRldmljZUJhY2tCdXR0b24oZXZlbnQpIHtcbiAgICAgICAgaWYgKHRoaXMucGFnZVN0YWNrLmxlbmd0aCA+IDEpIHtcbiAgICAgICAgICB0aGlzLnBvcFBhZ2UoKTtcbiAgICAgICAgfSBlbHNlIHtcbiAgICAgICAgICBldmVudC5jYWxsUGFyZW50SGFuZGxlcigpO1xuICAgICAgICB9XG4gICAgICB9LFxuICAgICAgX2ZpbmRTY3JvbGxQYWdlKHBhZ2UpIHtcbiAgICAgICAgY29uc3QgbmV4dFBhZ2UgPSBwYWdlLl9jb250ZW50RWxlbWVudC5jaGlsZHJlbi5sZW5ndGggPT09IDFcbiAgICAgICAgICAmJiB0aGlzLiRvbnMuX29ucy5fdXRpbC5nZXRUb3BQYWdlKHBhZ2UuX2NvbnRlbnRFbGVtZW50LmNoaWxkcmVuWzBdKTtcbiAgICAgICAgcmV0dXJuIG5leHRQYWdlID8gdGhpcy5fZmluZFNjcm9sbFBhZ2UobmV4dFBhZ2UpIDogcGFnZTtcbiAgICAgIH0sXG4gICAgICBfZWFjaFBhZ2Uoc3RhcnQsIGVuZCwgY2IpIHtcbiAgICAgICAgZm9yIChsZXQgaSA9IHN0YXJ0OyBpIDwgZW5kOyBpKyspIHtcbiAgICAgICAgICBjYih0aGlzLiRjaGlsZHJlbltpXS4kZWwpO1xuICAgICAgICB9XG4gICAgICB9LFxuICAgICAgX3JlYXR0YWNoUGFnZShwYWdlRWxlbWVudCwgcG9zaXRpb24gPSBudWxsLCByZXN0b3JlU2Nyb2xsKSB7XG4gICAgICAgIHRoaXMuJGVsLmluc2VydEJlZm9yZShwYWdlRWxlbWVudCwgcG9zaXRpb24pO1xuICAgICAgICByZXN0b3JlU2Nyb2xsIGluc3RhbmNlb2YgRnVuY3Rpb24gJiYgcmVzdG9yZVNjcm9sbCgpO1xuICAgICAgICBwYWdlRWxlbWVudC5faXNTaG93biA9IHRydWU7XG4gICAgICB9LFxuICAgICAgX3JlZGV0YWNoUGFnZShwYWdlRWxlbWVudCkge1xuICAgICAgICBwYWdlRWxlbWVudC5fZGVzdHJveSgpO1xuICAgICAgICByZXR1cm4gUHJvbWlzZS5yZXNvbHZlKCk7XG4gICAgICB9LFxuICAgICAgX2FuaW1hdGUoeyBsYXN0TGVuZ3RoLCBjdXJyZW50TGVuZ3RoLCBsYXN0VG9wUGFnZSwgY3VycmVudFRvcFBhZ2UsIHJlc3RvcmVTY3JvbGwgfSkge1xuICAgICAgICBjb25zdCBwdXNoZWRPcHRpb25zID0gdGhpcy5wYWdlU3RhY2tbdGhpcy5wYWdlU3RhY2subGVuZ3RoIC0gMV0ub25zTmF2aWdhdG9yT3B0aW9uc1xuICAgICAgICAgIHx8IGN1cnJlbnRUb3BQYWdlLl9fdnVlX18ub25zTmF2aWdhdG9yT3B0aW9uc1xuICAgICAgICAgIHx8IHt9O1xuXG4gICAgICAgIC8vIFB1c2hcbiAgICAgICAgaWYgKGN1cnJlbnRMZW5ndGggPiBsYXN0TGVuZ3RoKSB7XG4gICAgICAgICAgbGV0IGlzUmVhdHRhY2hlZCA9IGZhbHNlO1xuICAgICAgICAgIGlmIChsYXN0VG9wUGFnZS5wYXJlbnRFbGVtZW50ICE9PSB0aGlzLiRlbCkge1xuICAgICAgICAgICAgdGhpcy5fcmVhdHRhY2hQYWdlKGxhc3RUb3BQYWdlLCB0aGlzLiRlbC5jaGlsZHJlbltsYXN0TGVuZ3RoIC0gMV0sIHJlc3RvcmVTY3JvbGwpO1xuICAgICAgICAgICAgaXNSZWF0dGFjaGVkID0gdHJ1ZTtcbiAgICAgICAgICAgIGxhc3RMZW5ndGgtLTtcbiAgICAgICAgICB9XG5cbiAgICAgICAgICB0aGlzLl9lYWNoUGFnZShsYXN0TGVuZ3RoLCBjdXJyZW50TGVuZ3RoLCBlbCA9PiB7IGVsLnN0eWxlLnZpc2liaWxpdHkgPSAnaGlkZGVuJyB9KTtcbiAgICAgICAgICB0aGlzLl9lYWNoUGFnZShsYXN0TGVuZ3RoLCBjdXJyZW50TGVuZ3RoIC0gMSwgZWwgPT4geyBlbC5wdXNoZWRPcHRpb25zID0gcHVzaGVkT3B0aW9ucyB9KTtcblxuICAgICAgICAgIHJldHVybiB0aGlzLiRlbC5fcHVzaFBhZ2UoeyAuLi5wdXNoZWRPcHRpb25zLCBsZWF2ZVBhZ2U6IGxhc3RUb3BQYWdlIH0pXG4gICAgICAgICAgICAudGhlbigoKSA9PiB7XG4gICAgICAgICAgICAgIHNldEltbWVkaWF0ZSgoKSA9PiB7XG4gICAgICAgICAgICAgICAgdGhpcy5fZWFjaFBhZ2UobGFzdExlbmd0aCwgY3VycmVudExlbmd0aCwgZWwgPT4geyBlbC5zdHlsZS52aXNpYmlsaXR5ID0gJycgfSk7XG4gICAgICAgICAgICAgICAgdGhpcy5fZWFjaFBhZ2UobGFzdExlbmd0aCAtIDEsIGN1cnJlbnRMZW5ndGggLSAxLCBlbCA9PiB7IGVsLnN0eWxlLmRpc3BsYXkgPSAnbm9uZScgfSk7XG4gICAgICAgICAgICAgIH0pO1xuXG4gICAgICAgICAgICAgIGlmIChpc1JlYXR0YWNoZWQpIHtcbiAgICAgICAgICAgICAgICB0aGlzLl9yZWRldGFjaFBhZ2UobGFzdFRvcFBhZ2UpO1xuICAgICAgICAgICAgICB9XG4gICAgICAgICAgICB9KTtcbiAgICAgICAgfVxuXG4gICAgICAgIC8vIFBvcFxuICAgICAgICBpZiAoY3VycmVudExlbmd0aCA8IGxhc3RMZW5ndGgpIHtcbiAgICAgICAgICB0aGlzLl9yZWF0dGFjaFBhZ2UobGFzdFRvcFBhZ2UsIG51bGwsIHJlc3RvcmVTY3JvbGwpO1xuICAgICAgICAgIHJldHVybiB0aGlzLiRlbC5fcG9wUGFnZSh7IH0sICgpID0+IHRoaXMuX3JlZGV0YWNoUGFnZShsYXN0VG9wUGFnZSkpO1xuICAgICAgICB9XG5cbiAgICAgICAgLy8gUmVwbGFjZSBwYWdlXG4gICAgICAgIGN1cnJlbnRUb3BQYWdlLnN0eWxlLnZpc2liaWxpdHkgPSAnaGlkZGVuJztcbiAgICAgICAgdGhpcy5fcmVhdHRhY2hQYWdlKGxhc3RUb3BQYWdlLCBjdXJyZW50VG9wUGFnZSwgcmVzdG9yZVNjcm9sbCk7XG4gICAgICAgIHJldHVybiB0aGlzLiRlbC5fcHVzaFBhZ2UoeyAuLi5wdXNoZWRPcHRpb25zLCBfcmVwbGFjZVBhZ2U6IHRydWUgfSlcbiAgICAgICAgICAudGhlbigoKSA9PiB0aGlzLl9yZWRldGFjaFBhZ2UobGFzdFRvcFBhZ2UpKTtcbiAgICAgIH0sXG4gICAgICBfY2hlY2tTd2lwZShldmVudCkge1xuICAgICAgICBpZiAodGhpcy4kZWwuaGFzQXR0cmlidXRlKCdzd2lwZWFibGUnKSAmJlxuICAgICAgICAgIGV2ZW50LmxlYXZlUGFnZSAhPT0gdGhpcy4kZWwubGFzdENoaWxkICYmIGV2ZW50LmxlYXZlUGFnZSA9PT0gdGhpcy4kY2hpbGRyZW5bdGhpcy4kY2hpbGRyZW4ubGVuZ3RoIC0gMV0uJGVsXG4gICAgICAgICkge1xuICAgICAgICAgIHRoaXMucG9wUGFnZSgpO1xuICAgICAgICB9XG4gICAgICB9XG4gICAgfSxcblxuICAgIHdhdGNoOiB7XG4gICAgICBwYWdlU3RhY2soYWZ0ZXIsIGJlZm9yZSkge1xuICAgICAgICBpZiAodGhpcy4kZWwuaGFzQXR0cmlidXRlKCdzd2lwZWFibGUnKSAmJiB0aGlzLiRjaGlsZHJlbi5sZW5ndGggIT09IHRoaXMuJGVsLmNoaWxkcmVuLmxlbmd0aCkge1xuICAgICAgICAgIHJldHVybjtcbiAgICAgICAgfVxuXG4gICAgICAgIGNvbnN0IHByb3BXYXNNdXRhdGVkID0gYWZ0ZXIgPT09IGJlZm9yZTsgLy8gQ2FuIGJlIG11dGF0ZWQgb3IgcmVwbGFjZWRcbiAgICAgICAgY29uc3QgbGFzdFRvcFBhZ2UgPSB0aGlzLiRjaGlsZHJlblt0aGlzLiRjaGlsZHJlbi5sZW5ndGggLSAxXS4kZWw7XG4gICAgICAgIGNvbnN0IHNjcm9sbEVsZW1lbnQgPSB0aGlzLl9maW5kU2Nyb2xsUGFnZShsYXN0VG9wUGFnZSk7XG4gICAgICAgIGNvbnN0IHNjcm9sbFZhbHVlID0gc2Nyb2xsRWxlbWVudC5zY3JvbGxUb3AgfHwgMDtcblxuICAgICAgICB0aGlzLl9wYWdlU3RhY2tVcGRhdGUgPSB7XG4gICAgICAgICAgbGFzdFRvcFBhZ2UsXG4gICAgICAgICAgbGFzdExlbmd0aDogcHJvcFdhc011dGF0ZWQgPyB0aGlzLiRjaGlsZHJlbi5sZW5ndGggOiBiZWZvcmUubGVuZ3RoLFxuICAgICAgICAgIGN1cnJlbnRMZW5ndGg6ICFwcm9wV2FzTXV0YXRlZCAmJiBhZnRlci5sZW5ndGgsXG4gICAgICAgICAgcmVzdG9yZVNjcm9sbDogKCkgPT4gc2Nyb2xsRWxlbWVudC5zY3JvbGxUb3AgPSBzY3JvbGxWYWx1ZVxuICAgICAgICB9O1xuXG4gICAgICAgIC8vIHRoaXMuJG5leHRUaWNrKCgpID0+IHsgfSk7IC8vIFdhaXRzIHRvbyBsb25nLCB1cGRhdGVkKCkgaG9vayBpcyBmYXN0ZXIgYW5kIHByZXZlbnRzIGZsaWNrZXJpbmdzXG4gICAgICB9XG4gICAgfSxcblxuICAgIHVwZGF0ZWQoKSB7XG4gICAgICBpZiAodGhpcy5fcGFnZVN0YWNrVXBkYXRlKSB7XG4gICAgICAgIGxldCBjdXJyZW50VG9wUGFnZSA9IHRoaXMuJGNoaWxkcmVuW3RoaXMuJGNoaWxkcmVuLmxlbmd0aCAtIDFdLiRlbDtcbiAgICAgICAgbGV0IHsgbGFzdFRvcFBhZ2UsIGN1cnJlbnRMZW5ndGggfSA9IHRoaXMuX3BhZ2VTdGFja1VwZGF0ZTtcbiAgICAgICAgY29uc3QgeyBsYXN0TGVuZ3RoLCByZXN0b3JlU2Nyb2xsIH0gPSB0aGlzLl9wYWdlU3RhY2tVcGRhdGU7XG4gICAgICAgIGN1cnJlbnRMZW5ndGggPSBjdXJyZW50TGVuZ3RoID09PSBmYWxzZSA/IHRoaXMuJGNoaWxkcmVuLmxlbmd0aCA6IGN1cnJlbnRMZW5ndGg7XG5cbiAgICAgICAgaWYgKGN1cnJlbnRUb3BQYWdlICE9PSBsYXN0VG9wUGFnZSkge1xuICAgICAgICAgIHRoaXMuX3JlYWR5ID0gdGhpcy5fYW5pbWF0ZSh7IGxhc3RMZW5ndGgsIGN1cnJlbnRMZW5ndGgsIGxhc3RUb3BQYWdlLCBjdXJyZW50VG9wUGFnZSwgcmVzdG9yZVNjcm9sbCB9KTtcbiAgICAgICAgfSBlbHNlIGlmIChjdXJyZW50TGVuZ3RoICE9PSBsYXN0TGVuZ3RoKSB7XG4gICAgICAgICAgY3VycmVudFRvcFBhZ2UudXBkYXRlQmFja0J1dHRvbihjdXJyZW50TGVuZ3RoID4gMSk7XG4gICAgICAgIH1cblxuICAgICAgICBsYXN0VG9wUGFnZSA9IGN1cnJlbnRUb3BQYWdlID0gdGhpcy5fcGFnZVN0YWNrVXBkYXRlID0gbnVsbDtcbiAgICAgIH1cbiAgICB9XG4gIH07XG48L3NjcmlwdD5cbiIsIjx0ZW1wbGF0ZT5cbiAgPG9ucy1zcGxpdHRlci1zaWRlIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIj5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXNwbGl0dGVyLXNpZGU+XG48L3RlbXBsYXRlPlxuXG48c2NyaXB0PlxuICBpbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1zcGxpdHRlci1zaWRlJztcbiAgaW1wb3J0IHsgaGFzT3B0aW9ucywgZGVyaXZlRXZlbnRzIH0gZnJvbSAnLi4vbWl4aW5zJztcblxuICBleHBvcnQgZGVmYXVsdCB7XG4gICAgbmFtZTogJ3Ytb25zLXNwbGl0dGVyLXNpZGUnLFxuICAgIG1peGluczogW2hhc09wdGlvbnMsIGRlcml2ZUV2ZW50c10sXG5cbiAgICBwcm9wczoge1xuICAgICAgb3Blbjoge1xuICAgICAgICB0eXBlOiBCb29sZWFuLFxuICAgICAgICBkZWZhdWx0OiB1bmRlZmluZWRcbiAgICAgIH1cbiAgICB9LFxuXG4gICAgbWV0aG9kczoge1xuICAgICAgYWN0aW9uKCkge1xuICAgICAgICB0aGlzLl9zaG91bGRVcGRhdGUoKSAmJiB0aGlzLiRlbFt0aGlzLm9wZW4gPyAnb3BlbicgOiAnY2xvc2UnXS5jYWxsKHRoaXMuJGVsLCB0aGlzLm9wdGlvbnMpLmNhdGNoKCgpID0+IHt9KTtcbiAgICAgIH0sXG4gICAgICBfc2hvdWxkVXBkYXRlKCkge1xuICAgICAgICByZXR1cm4gdGhpcy5vcGVuICE9PSB1bmRlZmluZWQgJiYgdGhpcy5vcGVuICE9PSB0aGlzLiRlbC5pc09wZW47XG4gICAgICB9XG4gICAgfSxcblxuICAgIHdhdGNoOiB7XG4gICAgICBvcGVuKCkge1xuICAgICAgICB0aGlzLmFjdGlvbigpO1xuICAgICAgfVxuICAgIH0sXG5cbiAgICBtb3VudGVkKCkge1xuICAgICAgdGhpcy4kb24oWydwb3N0b3BlbicsICdwb3N0Y2xvc2UnLCAnbW9kZWNoYW5nZSddLCAoKSA9PiB0aGlzLl9zaG91bGRVcGRhdGUoKSAmJiB0aGlzLiRlbWl0KCd1cGRhdGU6b3BlbicsIHRoaXMuJGVsLmlzT3BlbikpO1xuXG4gICAgICB0aGlzLmFjdGlvbigpO1xuICAgIH1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8IS0tIFRoaXMgZWxlbWVudCBpcyB1c2VsZXNzIGV4Y2VwdCBmb3IgdGhlIGRlc3Ryb3kgcGFydCAtLT5cbiAgPG9ucy1sYXp5LXJlcGVhdD48L29ucy1sYXp5LXJlcGVhdD5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG5pbXBvcnQgJ29uc2VudWkvZXNtL2VsZW1lbnRzL29ucy1sYXp5LXJlcGVhdCc7XG5cbmV4cG9ydCBkZWZhdWx0IHtcbiAgbmFtZTogJ3Ytb25zLWxhenktcmVwZWF0JyxcblxuICBwcm9wczoge1xuICAgIHJlbmRlckl0ZW06IHtcbiAgICAgIHR5cGU6IEZ1bmN0aW9uLFxuICAgICAgcmVxdWlyZWQ6IHRydWUsXG4gICAgICB2YWxpZGF0b3IodmFsdWUpIHtcbiAgICAgICAgY29uc3QgY29tcG9uZW50ID0gdmFsdWUoMCk7XG4gICAgICAgIGlmIChjb21wb25lbnQuX2lzVnVlICYmICFjb21wb25lbnQuX2lzTW91bnRlZCkge1xuICAgICAgICAgIGNvbXBvbmVudC4kZGVzdHJveSgpO1xuICAgICAgICAgIHJldHVybiB0cnVlO1xuICAgICAgICB9XG4gICAgICAgIHJldHVybiBmYWxzZTtcbiAgICAgIH1cbiAgICB9LFxuICAgIGxlbmd0aDoge1xuICAgICAgdHlwZTogTnVtYmVyLFxuICAgICAgcmVxdWlyZWQ6IHRydWVcbiAgICB9LFxuICAgIGNhbGN1bGF0ZUl0ZW1IZWlnaHQ6IHtcbiAgICAgIHR5cGU6IEZ1bmN0aW9uLFxuICAgICAgZGVmYXVsdDogdW5kZWZpbmVkXG4gICAgfVxuICB9LFxuXG4gIGRhdGEoKSB7XG4gICAgcmV0dXJuIHtcbiAgICAgIHByb3ZpZGVyOiBudWxsXG4gICAgfTtcbiAgfSxcblxuICBtZXRob2RzOiB7XG4gICAgX3NldHVwKCkge1xuICAgICAgdGhpcy5wcm92aWRlciAmJiB0aGlzLnByb3ZpZGVyLmRlc3Ryb3koKTtcblxuICAgICAgY29uc3QgZGVsZWdhdGUgPSBuZXcgdGhpcy4kb25zLl9vbnMuX2ludGVybmFsLkxhenlSZXBlYXREZWxlZ2F0ZSh7XG4gICAgICAgIGNhbGN1bGF0ZUl0ZW1IZWlnaHQ6IHRoaXMuY2FsY3VsYXRlSXRlbUhlaWdodCxcbiAgICAgICAgY3JlYXRlSXRlbUNvbnRlbnQ6IGkgPT4gdGhpcy5yZW5kZXJJdGVtKGkpLiRtb3VudCgpLiRlbCxcbiAgICAgICAgZGVzdHJveUl0ZW06IChpLCB7IGVsZW1lbnQgfSkgPT4gZWxlbWVudC5fX3Z1ZV9fLiRkZXN0cm95KCksXG4gICAgICAgIGNvdW50SXRlbXM6ICgpID0+IHRoaXMubGVuZ3RoXG4gICAgICB9LCBudWxsKTtcblxuICAgICAgdGhpcy5wcm92aWRlciA9IG5ldyB0aGlzLiRvbnMuX29ucy5faW50ZXJuYWwuTGF6eVJlcGVhdFByb3ZpZGVyKHRoaXMuJHBhcmVudC4kZWwsIGRlbGVnYXRlKTtcbiAgICB9LFxuICAgIHJlZnJlc2goKSB7XG4gICAgICByZXR1cm4gdGhpcy5wcm92aWRlci5yZWZyZXNoKCk7XG4gICAgfVxuICB9LFxuXG4gIHdhdGNoOiB7XG4gICAgcmVuZGVySXRlbSgpIHtcbiAgICAgIHRoaXMuX3NldHVwKCk7XG4gICAgfSxcbiAgICBsZW5ndGgoKSB7XG4gICAgICB0aGlzLl9zZXR1cCgpO1xuICAgIH0sXG4gICAgY2FsY3VsYXRlSXRlbUhlaWdodCgpIHtcbiAgICAgIHRoaXMuX3NldHVwKCk7XG4gICAgfVxuICB9LFxuXG4gIG1vdW50ZWQoKSB7XG4gICAgdGhpcy5fc2V0dXAoKTtcbiAgICB0aGlzLiR2bm9kZS5jb250ZXh0LiRvbigncmVmcmVzaCcsIHRoaXMucmVmcmVzaCk7XG4gIH0sXG5cbiAgYmVmb3JlRGVzdHJveSgpIHtcbiAgICB0aGlzLiR2bm9kZS5jb250ZXh0LiRvZmYoJ3JlZnJlc2gnLCB0aGlzLnJlZnJlc2gpO1xuXG4gICAgLy8gVGhpcyB3aWxsIGRlc3Ryb3kgdGhlIHByb3ZpZGVyIG9uY2UgdGhlIHJlbmRlcmVkIGVsZW1lbnRcbiAgICAvLyBpcyBkZXRhY2hlZCAoZGV0YWNoZWRDYWxsYmFjaykuIFRoZXJlZm9yZSwgYW5pbWF0aW9uc1xuICAgIC8vIGhhdmUgdGltZSB0byBmaW5pc2ggYmVmb3JlIGVsZW1lbnRzIHN0YXJ0IHRvIGRpc2FwcGVhci5cbiAgICAvLyBJdCBjYW5ub3QgYmUgc2V0IGVhcmxpZXIgaW4gb3JkZXIgdG8gcHJldmVudCBhY2NpZGVudGFsXG4gICAgLy8gZGVzdHJveXMgaWYgdGhpcyBlbGVtZW50IGlzIHJldGFjaGVkIGJ5IHNvbWV0aGluZyBlbHNlLlxuICAgIHRoaXMuJGVsLl9sYXp5UmVwZWF0UHJvdmlkZXIgPSB0aGlzLnByb3ZpZGVyO1xuICAgIHRoaXMucHJvdmlkZXIgPSBudWxsO1xuICB9XG59O1xuPC9zY3JpcHQ+XG4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtc2VsZWN0IHYtb249XCIkbGlzdGVuZXJzXCIgOm1vZGlmaWVyPVwibm9ybWFsaXplZE1vZGlmaWVyXCI+XG4gICAgPHNlbGVjdD5cbiAgICAgIDxzbG90Pjwvc2xvdD5cbiAgICA8L3NlbGVjdD5cbiAgPC9vbnMtc2VsZWN0PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtc2VsZWN0JztcbiAgaW1wb3J0IHsgbW9kZWxJbnB1dCwgbW9kaWZpZXIgfSBmcm9tICcuLi9taXhpbnMnO1xuXG4gIGV4cG9ydCBkZWZhdWx0IHtcbiAgICBuYW1lOiAndi1vbnMtc2VsZWN0JyxcbiAgICBtaXhpbnM6IFttb2RlbElucHV0LCBtb2RpZmllcl1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8b25zLXNlZ21lbnQgOmFjdGl2ZS1pbmRleD1cImluZGV4XCIgQHBvc3RjaGFuZ2Uuc2VsZj1cIiRlbWl0KCd1cGRhdGU6aW5kZXgnLCAkZXZlbnQuaW5kZXgpXCI+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1zZWdtZW50PlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtc2VnbWVudCc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1zZWdtZW50JyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHNdLFxuXG4gICAgcHJvcHM6IHtcbiAgICAgIGluZGV4OiB7XG4gICAgICAgIHR5cGU6IE51bWJlclxuICAgICAgfVxuICAgIH0sXG5cbiAgICB3YXRjaDoge1xuICAgICAgaW5kZXgoKSB7XG4gICAgICAgIGlmICh0aGlzLmluZGV4ICE9PSB0aGlzLiRlbC5nZXRBY3RpdmVCdXR0b25JbmRleCgpKSB7XG4gICAgICAgICAgdGhpcy4kZWwuc2V0QWN0aXZlQnV0dG9uKHRoaXMuaW5kZXgsIHsgcmVqZWN0OiBmYWxzZSB9KTtcbiAgICAgICAgfVxuICAgICAgfVxuICAgIH1cbiAgfTtcbjwvc2NyaXB0PlxuIiwiPHRlbXBsYXRlPlxuICA8b25zLXB1bGwtaG9va1xuICAgIDpvbi1hY3Rpb24ucHJvcD1cImFjdGlvblwiXG4gICAgOm9uLXB1bGwucHJvcD1cIm9uUHVsbFwiXG4gICAgdi1vbj1cInVucmVjb2duaXplZExpc3RlbmVyc1wiXG4gID5cbiAgICA8c2xvdD48L3Nsb3Q+XG4gIDwvb25zLXB1bGwtaG9vaz5cbjwvdGVtcGxhdGU+XG5cbjxzY3JpcHQ+XG4gIGltcG9ydCAnb25zZW51aS9lc20vZWxlbWVudHMvb25zLXB1bGwtaG9vayc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cyB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1wdWxsLWhvb2snLFxuICAgIG1peGluczogW2Rlcml2ZUV2ZW50c10sXG5cbiAgICBwcm9wczoge1xuICAgICAgYWN0aW9uOiB7XG4gICAgICAgIHR5cGU6IEZ1bmN0aW9uXG4gICAgICB9LFxuICAgICAgb25QdWxsOiB7XG4gICAgICAgIHR5cGU6IEZ1bmN0aW9uXG4gICAgICB9XG4gICAgfVxuICB9O1xuPC9zY3JpcHQ+XG4iLCI8dGVtcGxhdGU+XG4gIDxvbnMtcGFnZVxuICAgIDpvbi1pbmZpbml0ZS1zY3JvbGwucHJvcD1cImluZmluaXRlU2Nyb2xsXCJcbiAgICA6bW9kaWZpZXI9XCJub3JtYWxpemVkTW9kaWZpZXJcIlxuICAgIHYtb249XCJ1bnJlY29nbml6ZWRMaXN0ZW5lcnNcIlxuICA+XG4gICAgPHNsb3Q+PC9zbG90PlxuICA8L29ucy1wYWdlPlxuPC90ZW1wbGF0ZT5cblxuPHNjcmlwdD5cbiAgaW1wb3J0ICdvbnNlbnVpL2VzbS9lbGVtZW50cy9vbnMtcGFnZSc7XG4gIGltcG9ydCB7IGRlcml2ZUV2ZW50cywgZGVyaXZlREJCLCBtb2RpZmllciB9IGZyb20gJy4uL21peGlucyc7XG5cbiAgZXhwb3J0IGRlZmF1bHQge1xuICAgIG5hbWU6ICd2LW9ucy1wYWdlJyxcbiAgICBtaXhpbnM6IFtkZXJpdmVFdmVudHMsIGRlcml2ZURCQiwgbW9kaWZpZXJdLFxuXG4gICAgcHJvcHM6IHtcbiAgICAgIGluZmluaXRlU2Nyb2xsOiB7XG4gICAgICAgIHR5cGU6IEZ1bmN0aW9uXG4gICAgICB9XG4gICAgfVxuICB9O1xuPC9zY3JpcHQ+XG4iLCIvLyBHZW5lcmljIGNvbXBvbmVudHM6XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNUb29sYmFyIH0gZnJvbSAnLi9WT25zVG9vbGJhcic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNCb3R0b21Ub29sYmFyIH0gZnJvbSAnLi9WT25zQm90dG9tVG9vbGJhcic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNUb29sYmFyQnV0dG9uIH0gZnJvbSAnLi9WT25zVG9vbGJhckJ1dHRvbic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNBbGVydERpYWxvZ0J1dHRvbiB9IGZyb20gJy4vVk9uc0FsZXJ0RGlhbG9nQnV0dG9uJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0J1dHRvbiB9IGZyb20gJy4vVk9uc0J1dHRvbic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNJY29uIH0gZnJvbSAnLi9WT25zSWNvbic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNDYXJkIH0gZnJvbSAnLi9WT25zQ2FyZCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNMaXN0IH0gZnJvbSAnLi9WT25zTGlzdCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNMaXN0SXRlbSB9IGZyb20gJy4vVk9uc0xpc3RJdGVtJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0xpc3RUaXRsZSB9IGZyb20gJy4vVk9uc0xpc3RUaXRsZSc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNMaXN0SGVhZGVyIH0gZnJvbSAnLi9WT25zTGlzdEhlYWRlcic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNSaXBwbGUgfSBmcm9tICcuL1ZPbnNSaXBwbGUnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zUm93IH0gZnJvbSAnLi9WT25zUm93JztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0NvbCB9IGZyb20gJy4vVk9uc0NvbCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNQcm9ncmVzc0JhciB9IGZyb20gJy4vVk9uc1Byb2dyZXNzQmFyJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1Byb2dyZXNzQ2lyY3VsYXIgfSBmcm9tICcuL1ZPbnNQcm9ncmVzc0NpcmN1bGFyJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0Nhcm91c2VsSXRlbSB9IGZyb20gJy4vVk9uc0Nhcm91c2VsSXRlbSc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNTcGxpdHRlck1hc2sgfSBmcm9tICcuL1ZPbnNTcGxpdHRlck1hc2snO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zU3BsaXR0ZXJDb250ZW50IH0gZnJvbSAnLi9WT25zU3BsaXR0ZXJDb250ZW50JztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1NwbGl0dGVyIH0gZnJvbSAnLi9WT25zU3BsaXR0ZXInO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zU3dpdGNoIH0gZnJvbSAnLi9WT25zU3dpdGNoJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0NoZWNrYm94IH0gZnJvbSAnLi9WT25zQ2hlY2tib3gnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zSW5wdXQgfSBmcm9tICcuL1ZPbnNJbnB1dCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNTZWFyY2hJbnB1dCB9IGZyb20gJy4vVk9uc1NlYXJjaElucHV0JztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1JhbmdlIH0gZnJvbSAnLi9WT25zUmFuZ2UnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zUmFkaW8gfSBmcm9tICcuL1ZPbnNSYWRpbyc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNGYWIgfSBmcm9tICcuL1ZPbnNGYWInO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zU3BlZWREaWFsSXRlbSB9IGZyb20gJy4vVk9uc1NwZWVkRGlhbEl0ZW0nO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zRGlhbG9nIH0gZnJvbSAnLi9WT25zRGlhbG9nJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0FjdGlvblNoZWV0IH0gZnJvbSAnLi9WT25zQWN0aW9uU2hlZXQnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zQWN0aW9uU2hlZXRCdXR0b24gfSBmcm9tICcuL1ZPbnNBY3Rpb25TaGVldEJ1dHRvbic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNNb2RhbCB9IGZyb20gJy4vVk9uc01vZGFsJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1RvYXN0IH0gZnJvbSAnLi9WT25zVG9hc3QnO1xuXG4vLyBNYW51YWwgY29tcG9uZW50czpcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1BvcG92ZXIgfSBmcm9tICcuL1ZPbnNQb3BvdmVyJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0FsZXJ0RGlhbG9nIH0gZnJvbSAnLi9WT25zQWxlcnREaWFsb2cnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zU3BlZWREaWFsIH0gZnJvbSAnLi9WT25zU3BlZWREaWFsJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0Nhcm91c2VsIH0gZnJvbSAnLi9WT25zQ2Fyb3VzZWwnO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zVGFiIH0gZnJvbSAnLi9WT25zVGFiJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1RhYmJhciB9IGZyb20gJy4vVk9uc1RhYmJhcic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNCYWNrQnV0dG9uIH0gZnJvbSAnLi9WT25zQmFja0J1dHRvbic7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNOYXZpZ2F0b3IgfSBmcm9tICcuL1ZPbnNOYXZpZ2F0b3InO1xuZXhwb3J0IHsgZGVmYXVsdCBhcyBWT25zU3BsaXR0ZXJTaWRlIH0gZnJvbSAnLi9WT25zU3BsaXR0ZXJTaWRlJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc0xhenlSZXBlYXQgfSBmcm9tICcuL1ZPbnNMYXp5UmVwZWF0JztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1NlbGVjdCB9IGZyb20gJy4vVk9uc1NlbGVjdCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNTZWdtZW50IH0gZnJvbSAnLi9WT25zU2VnbWVudCc7XG5leHBvcnQgeyBkZWZhdWx0IGFzIFZPbnNQdWxsSG9vayB9IGZyb20gJy4vVk9uc1B1bGxIb29rJztcbmV4cG9ydCB7IGRlZmF1bHQgYXMgVk9uc1BhZ2UgfSBmcm9tICcuL1ZPbnNQYWdlJztcbiIsImltcG9ydCBvbnMgZnJvbSAnb25zZW51aSc7XG5pbXBvcnQgc2V0dXAgZnJvbSAnLi9zZXR1cCc7XG5pbXBvcnQgKiBhcyBjb21wb25lbnRzIGZyb20gJy4vY29tcG9uZW50cyc7XG5cbmNvbnN0ICRvbnMgPSBzZXR1cChvbnMpO1xuXG4kb25zLmluc3RhbGwgPSAoVnVlLCBwYXJhbXMgPSB7fSkgPT4ge1xuICAvKipcbiAgICogUmVnaXN0ZXIgY29tcG9uZW50cyBvZiB2dWUtb25zZW51aS5cbiAgICovXG4gIE9iamVjdC5rZXlzKGNvbXBvbmVudHMpXG4gICAgLmZvckVhY2goa2V5ID0+IFZ1ZS5jb21wb25lbnQoY29tcG9uZW50c1trZXldLm5hbWUsIGNvbXBvbmVudHNba2V5XSkpO1xuXG4gIC8qKlxuICAgKiBFeHBvc2Ugb25zIG9iamVjdC5cbiAgICovXG4gIFZ1ZS5wcm90b3R5cGUuJG9ucyA9ICRvbnM7XG59O1xuXG5pZiAodHlwZW9mIHdpbmRvdyAhPT0gJ3VuZGVmaW5lZCcgJiYgd2luZG93LlZ1ZSkge1xuICB3aW5kb3cuVnVlLnVzZSh7IGluc3RhbGw6ICRvbnMuaW5zdGFsbCB9KTtcbn1cblxuZXhwb3J0IGRlZmF1bHQgJG9ucztcbiJdLCJuYW1lcyI6WyJvbnMiLCJPYmplY3QiLCJrZXlzIiwiZmlsdGVyIiwic29tZSIsImsiLCJtYXRjaCIsInQiLCJyZWR1Y2UiLCJyIiwiX29ucyIsImNhcGl0YWxpemUiLCJzdHJpbmciLCJjaGFyQXQiLCJ0b1VwcGVyQ2FzZSIsInNsaWNlIiwiY2FtZWxpemUiLCJ0b0xvd2VyQ2FzZSIsInJlcGxhY2UiLCJtIiwibCIsImV2ZW50VG9IYW5kbGVyIiwibmFtZSIsImhhbmRsZXJUb1Byb3AiLCJfc2V0dXBEQkIiLCJkYmIiLCJoYW5kbGVyIiwiY29tcG9uZW50IiwiJGVsIiwiX2NhbGxiYWNrIiwiZSIsImNhbGxQYXJlbnRIYW5kbGVyIiwicnVuRGVmYXVsdCIsIiRlbWl0IiwiZXZlbnQiLCJfaXNEQkJTZXR1cCIsImRlcml2ZURCQiIsIm9uRGV2aWNlQmFja0J1dHRvbiIsImRlc3Ryb3kiLCJkZXJpdmVFdmVudHMiLCIkb3B0aW9ucyIsIl9jb21wb25lbnRUYWciLCIkbGlzdGVuZXJzIiwiJG9ucyIsImVsZW1lbnRzIiwiZXZlbnRzIiwiaW5kZXhPZiIsIl9oYW5kbGVycyIsImNvbnN0cnVjdG9yIiwiZm9yRWFjaCIsImtleSIsInRhcmdldCIsInRlc3QiLCJ0YWdOYW1lIiwiYWRkRXZlbnRMaXN0ZW5lciIsInJlbW92ZUV2ZW50TGlzdGVuZXIiLCJfdG9nZ2xlVmlzaWJpbGl0eSIsInZpc2libGUiLCJjYWxsIiwibm9ybWFsaXplZE9wdGlvbnMiLCJvcHRpb25zIiwiX3RlbGVwb3J0IiwiX2lzRGVzdHJveWVkIiwicGFyZW50Tm9kZSIsImRvY3VtZW50IiwiYm9keSIsImFwcGVuZENoaWxkIiwiX3VubW91bnQiLCJoaWRlIiwidGhlbiIsInJlbW92ZSIsImhpZGFibGUiLCJCb29sZWFuIiwidW5kZWZpbmVkIiwiJG5leHRUaWNrIiwiaGFzT3B0aW9ucyIsInNlbGZQcm92aWRlciIsImRpYWxvZ0NhbmNlbCIsIiRvbiIsInBvcnRhbCIsIm1vZGlmaWVyIiwiU3RyaW5nIiwiQXJyYXkiLCJpc0FycmF5Iiwiam9pbiIsImFjYyIsInRyaW0iLCJtb2RlbCIsIm1vZGVsSW5wdXQiLCJwcm9wIiwiTnVtYmVyIiwidmFsdWUiLCJfdXBkYXRlVmFsdWUiLCJfb25Nb2RlbEV2ZW50IiwibW9kZWxDaGVja2JveCIsImNoZWNrZWQiLCJuZXdWYWx1ZSIsImluZGV4IiwiaW5jbHVkZWQiLCJsZW5ndGgiLCJtb2RlbFJhZGlvIiwicmVuZGVyIiwiYWN0aW9uIiwiZXhwYW5kZWQiLCJfaXNWdWUiLCJFdmVudCIsIkhUTUxFbGVtZW50Iiwibm9ybWFsaXplZFRhcmdldCIsImV2ZXJ5IiwiRnVuY3Rpb24iLCJwcmV2ZW50RGVmYXVsdCIsInRvZ2dsZUl0ZW1zIiwib3BlbiIsImlzT3BlbiIsIl9zaG91bGRVcGRhdGUiLCJfdXBkYXRlVG9nZ2xlIiwiZ2V0QWN0aXZlSW5kZXgiLCJzZXRBY3RpdmVJbmRleCIsInRhYmJhciIsInNldEFjdGl2ZVRhYiIsInJlamVjdCIsInNldEFjdGl2ZSIsImFjdGl2ZSIsImdldE93blByb3BlcnR5RGVzY3JpcHRvciIsInRhYiIsImxhYmVsIiwiaWNvbiIsImdldEFjdGl2ZVRhYkluZGV4IiwibmF2aWdhdG9yIiwicGFnZVN0YWNrIiwicG9wUGFnZSIsInBvcCIsImhhc093blByb3BlcnR5IiwiX3JlYWR5IiwiUHJvbWlzZSIsInJlc29sdmUiLCJwYWdlIiwibmV4dFBhZ2UiLCJfY29udGVudEVsZW1lbnQiLCJjaGlsZHJlbiIsIl91dGlsIiwiZ2V0VG9wUGFnZSIsIl9maW5kU2Nyb2xsUGFnZSIsInN0YXJ0IiwiZW5kIiwiY2IiLCJpIiwiJGNoaWxkcmVuIiwicGFnZUVsZW1lbnQiLCJwb3NpdGlvbiIsInJlc3RvcmVTY3JvbGwiLCJpbnNlcnRCZWZvcmUiLCJfaXNTaG93biIsIl9kZXN0cm95IiwibGFzdExlbmd0aCIsImN1cnJlbnRMZW5ndGgiLCJsYXN0VG9wUGFnZSIsImN1cnJlbnRUb3BQYWdlIiwicHVzaGVkT3B0aW9ucyIsIm9uc05hdmlnYXRvck9wdGlvbnMiLCJfX3Z1ZV9fIiwiaXNSZWF0dGFjaGVkIiwicGFyZW50RWxlbWVudCIsIl9yZWF0dGFjaFBhZ2UiLCJfZWFjaFBhZ2UiLCJzdHlsZSIsInZpc2liaWxpdHkiLCJfcHVzaFBhZ2UiLCJsZWF2ZVBhZ2UiLCJkaXNwbGF5IiwiX3JlZGV0YWNoUGFnZSIsIl9wb3BQYWdlIiwiX3JlcGxhY2VQYWdlIiwiaGFzQXR0cmlidXRlIiwibGFzdENoaWxkIiwiYWZ0ZXIiLCJiZWZvcmUiLCJwcm9wV2FzTXV0YXRlZCIsInNjcm9sbEVsZW1lbnQiLCJzY3JvbGxWYWx1ZSIsInNjcm9sbFRvcCIsIl9wYWdlU3RhY2tVcGRhdGUiLCJfYW5pbWF0ZSIsInVwZGF0ZUJhY2tCdXR0b24iLCJjYXRjaCIsIl9pc01vdW50ZWQiLCIkZGVzdHJveSIsInByb3ZpZGVyIiwiZGVsZWdhdGUiLCJfaW50ZXJuYWwiLCJMYXp5UmVwZWF0RGVsZWdhdGUiLCJjYWxjdWxhdGVJdGVtSGVpZ2h0IiwicmVuZGVySXRlbSIsIiRtb3VudCIsImVsZW1lbnQiLCJMYXp5UmVwZWF0UHJvdmlkZXIiLCIkcGFyZW50IiwicmVmcmVzaCIsIl9zZXR1cCIsIiR2bm9kZSIsImNvbnRleHQiLCIkb2ZmIiwiX2xhenlSZXBlYXRQcm92aWRlciIsImdldEFjdGl2ZUJ1dHRvbkluZGV4Iiwic2V0QWN0aXZlQnV0dG9uIiwic2V0dXAiLCJpbnN0YWxsIiwiVnVlIiwiY29tcG9uZW50cyIsInByb3RvdHlwZSIsIndpbmRvdyIsInVzZSJdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7OztBQUFBLFlBQWUsVUFBU0EsTUFBVCxFQUFjO1NBQ3BCQyxPQUFPQyxJQUFQLENBQVlGLE1BQVosRUFDSkcsTUFESSxDQUNHO1dBQUssQ0FDWCxLQURXLEVBRVgsVUFGVyxFQUdYLFNBSFcsRUFJWCxPQUpXLEVBS1gsT0FMVyxFQU1YLE1BTlcsRUFPWCxRQVBXLEVBUVgsVUFSVyxFQVNYLFdBVFcsRUFVWCxpQkFWVyxFQVdYLGNBWFcsRUFZWCxhQVpXLEVBYVgsVUFiVyxFQWNYLE9BZFcsRUFlWEMsSUFmVyxDQWVOO2FBQUtDLEVBQUVDLEtBQUYsQ0FBUUMsQ0FBUixDQUFMO0tBZk0sQ0FBTDtHQURILEVBaUJKQyxNQWpCSSxDQWlCRyxVQUFDQyxDQUFELEVBQUlKLENBQUosRUFBVTtNQUNkQSxDQUFGLElBQU9MLE9BQUlLLENBQUosQ0FBUDtXQUNPSSxDQUFQO0dBbkJHLEVBb0JGLEVBQUVDLE1BQU1WLE1BQVIsRUFwQkUsQ0FBUDs7O0FDQ0ssSUFBTVcsYUFBYSxTQUFiQSxVQUFhO1NBQVVDLE9BQU9DLE1BQVAsQ0FBYyxDQUFkLEVBQWlCQyxXQUFqQixLQUFpQ0YsT0FBT0csS0FBUCxDQUFhLENBQWIsQ0FBM0M7Q0FBbkI7O0FBRVAsQUFBTyxJQUFNQyxXQUFXLFNBQVhBLFFBQVc7U0FBVUosT0FBT0ssV0FBUCxHQUFxQkMsT0FBckIsQ0FBNkIsV0FBN0IsRUFBMEMsVUFBQ0MsQ0FBRCxFQUFJQyxDQUFKO1dBQVVBLEVBQUVOLFdBQUYsRUFBVjtHQUExQyxDQUFWO0NBQWpCOztBQUVQLEFBQU8sSUFBTU8saUJBQWlCLFNBQWpCQSxjQUFpQjtTQUFRLFFBQVFWLFdBQVdXLElBQVgsQ0FBaEI7Q0FB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

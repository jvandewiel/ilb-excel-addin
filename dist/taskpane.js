/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/services/dataLoaderService.ts":
/*!*******************************************!*\
  !*** ./src/services/dataLoaderService.ts ***!
  \*******************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   DataLoaderService: function() { return /* binding */ DataLoaderService; },
/* harmony export */   dataLoader: function() { return /* binding */ dataLoader; }
/* harmony export */ });
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function ownKeys(e, r) { var t = Object.keys(e); if (Object.getOwnPropertySymbols) { var o = Object.getOwnPropertySymbols(e); r && (o = o.filter(function (r) { return Object.getOwnPropertyDescriptor(e, r).enumerable; })), t.push.apply(t, o); } return t; }
function _objectSpread(e) { for (var r = 1; r < arguments.length; r++) { var t = null != arguments[r] ? arguments[r] : {}; r % 2 ? ownKeys(Object(t), !0).forEach(function (r) { _defineProperty(e, r, t[r]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function (r) { Object.defineProperty(e, r, Object.getOwnPropertyDescriptor(t, r)); }); } return e; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function _classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function _defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, _toPropertyKey(o.key), o); } }
function _createClass(e, r, t) { return r && _defineProperties(e.prototype, r), t && _defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
/*
 * Data Loader Service for Inlenersbeloning Excel Plugin
 * Loads remuneration data from local JSON file
 */

/**
 * Data loader service to load remuneration data from JSON file
 */
var DataLoaderService = /*#__PURE__*/function () {
  function DataLoaderService() {
    _classCallCheck(this, DataLoaderService);
    _defineProperty(this, "remunerationData", []);
    _defineProperty(this, "isLoaded", false);
  }
  return _createClass(DataLoaderService, [{
    key: "loadData",
    value: (
    /**
     * Load the remuneration data from JSON file
     */
    function () {
      var _loadData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
        var response, _t;
        return _regenerator().w(function (_context) {
          while (1) switch (_context.p = _context.n) {
            case 0:
              if (!this.isLoaded) {
                _context.n = 1;
                break;
              }
              return _context.a(2);
            case 1:
              _context.p = 1;
              console.log("Loading remuneration data from JSON file...");
              _context.n = 2;
              return fetch("./assets/renumerations-export.json");
            case 2:
              response = _context.v;
              if (response.ok) {
                _context.n = 3;
                break;
              }
              throw new Error("Failed to load data: ".concat(response.status, " ").concat(response.statusText));
            case 3:
              _context.n = 4;
              return response.json();
            case 4:
              this.remunerationData = _context.v;
              this.isLoaded = true;
              console.log("Loaded ".concat(this.remunerationData.length, " remuneration records"));
              _context.n = 6;
              break;
            case 5:
              _context.p = 5;
              _t = _context.v;
              console.error("Error loading remuneration data:", _t);
              throw new Error("Failed to load remuneration data: ".concat(_t.message));
            case 6:
              return _context.a(2);
          }
        }, _callee, this, [[1, 5]]);
      }));
      function loadData() {
        return _loadData.apply(this, arguments);
      }
      return loadData;
    }()
    /**
     * Get all unique partners from the data
     */
    )
  }, {
    key: "getPartners",
    value: function getPartners() {
      if (!this.isLoaded) {
        console.warn("Data not loaded yet. Call loadData() first.");
        return [];
      }
      var partnersMap = new Map();
      this.remunerationData.forEach(function (item) {
        if (item.company && item.company.id && item.company.businessName) {
          partnersMap.set(item.company.id, {
            id: item.company.id,
            businessName: item.company.businessName
          });
        }
      });
      var partners = Array.from(partnersMap.values());
      console.log("Found ".concat(partners.length, " unique partners"));
      return partners.sort(function (a, b) {
        return a.businessName.localeCompare(b.businessName);
      });
    }

    /**
     * Get remuneration data for a specific partner
     */
  }, {
    key: "getPartnerRemuneration",
    value: function getPartnerRemuneration(partnerId) {
      if (!this.isLoaded) {
        console.warn("Data not loaded yet. Call loadData() first.");
        return [];
      }
      var partnerData = this.remunerationData.filter(function (item) {
        return item.company && item.company.id === partnerId;
      });
      console.log("Found ".concat(partnerData.length, " remuneration records for partner: ").concat(partnerId));
      return partnerData;
    }

    /**
     * Utility function to format amount with percentage symbol if needed
     * @param amount The amount value
     * @param amountType Object containing amount type information
     * @returns Formatted amount string
     */
  }, {
    key: "formatAmountDisplay",
    value: function formatAmountDisplay(amount, amountType) {
      if (!amount) return "N/A";
      return "".concat(amount).concat((amountType === null || amountType === void 0 ? void 0 : amountType.value) === "Percentage" ? "%" : "");
    }

    /**
     * Maps pensionable earning code values to component properties
     * @param code The code value from pensionable earning
     * @returns The corresponding component name
     */
  }, {
    key: "mapCodeToComponent",
    value: function mapCodeToComponent(code) {
      switch (code) {
        case "CAOAllowances":
          return "allowances";
        case "CAOHolidayAllowance":
          return "holidayAllowance";
        case "CAOOneTimePayment":
          return "oneTimePayment";
        case "CAOFixedPayment":
          return "fixedPayment";
        case "CAOVariablePayment":
          return "variablePayment";
        case "CAOHoliday":
          return "holiday";
        case "CAOWageTable":
          return "wageTable";
        case "CAOWageIncrement":
          return "wageIncrement";
        case "CAOPeriodicIncrement":
          return "periodicIncrement";
        case "CAORetroactiveIncrement":
          return "retroactiveIncrement";
        case "Other":
          return "other";
        default:
          return "unknown";
      }
    }

    /**
     * Utility function to enhance pensionable earnings with component mapping
     * @param pensionableEarnings Array of pensionable earnings
     * @returns Array of pensionable earnings with component property added
     */
  }, {
    key: "enhancePensionableEarningsWithComponents",
    value: function enhancePensionableEarningsWithComponents(pensionableEarnings) {
      var _this = this;
      return pensionableEarnings.map(function (earning) {
        return _objectSpread(_objectSpread({}, earning), {}, {
          component: _this.mapCodeToComponent(earning.code)
        });
      });
    }

    /**
     * Generic function to generate Excel data with consistent structure
     * @param partnerId Partner ID to filter data
     * @param noDataMessage Message to show when no data is found
     * @param headers Array of column headers
     * @param dataProcessor Function that processes partner data and returns rows
     */
  }, {
    key: "generateDataWithStandardStructure",
    value: function generateDataWithStandardStructure(partnerId, noDataMessage, headers, dataProcessor) {
      var partnerData = this.getPartnerRemuneration(partnerId);
      var noDataFound = Array(headers.length).fill("").map(function (_, index) {
        return index === 0 ? noDataMessage : "";
      });
      if (!partnerData || partnerData.length === 0) {
        return [noDataFound];
      }
      var result = [];

      // Add header
      result.push(headers);

      // Process data using the provided processor function
      var dataRows = dataProcessor(partnerData);
      result.push.apply(result, _toConsumableArray(dataRows));

      // If no data was processed, add info message
      if (result.length === 1) {
        result.push(noDataFound);
      }
      return result;
    }

    /**
     * Generate metadata for Excel sheet (rows 1-10)
     */
  }, {
    key: "generateMetadata",
    value: function generateMetadata(partnerId) {
      var partnerData = this.getPartnerRemuneration(partnerId);
      if (!partnerData || partnerData.length === 0) {
        return [["Metadata", "", ""], ["Partner Name", "No data available", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""]];
      }
      var firstRecord = partnerData[0];
      var company = firstRecord.company;
      var workingHours = firstRecord.workingHours;
      return [["ILB Remuneration Data", "", ""], ["Partner Name", (company === null || company === void 0 ? void 0 : company.businessName) || "Unknown", ""], ["Commerce Number", (company === null || company === void 0 ? void 0 : company.commerceNumber) || "N/A", ""], ["Version", "".concat(firstRecord.lenderVersion, "/").concat(firstRecord.totalVersions), ""], ["Start Date", firstRecord.startDate || "N/A", ""], ["Publish Date", firstRecord.publishDate || "N/A", ""], ["Full Time Hours", (workingHours === null || workingHours === void 0 ? void 0 : workingHours.fullTimeHours) || "N/A", (workingHours === null || workingHours === void 0 ? void 0 : workingHours.fullTimeHoursPer) || ""], ["Working Hours Comments", (workingHours === null || workingHours === void 0 ? void 0 : workingHours.workingHoursComments) || "N/A", ""], ["", "", ""], ["Remuneration Components:", "", ""]];
    }

    /**
     * Generate holiday allowances data for Excel format
     * Takes holidayAllowances objects from remuneration data
     * Returns formatted data with holiday allowance details
     */
  }, {
    key: "generateHolidayAllowances",
    value: function generateHolidayAllowances(partnerId) {
      var _this2 = this;
      var partnerData = this.getPartnerRemuneration(partnerId);
      var noDataFound = ["No holiday allowances data available", "", "", "", "", "", "", "", ""];
      if (!partnerData || partnerData.length === 0) {
        return [noDataFound];
      }
      var result = [];

      // Check if holidayAllowance is pensionable for this partner
      var pensionableComponents = this.getPensionableComponents(partnerId);
      var isHolidayAllowancePensionable = pensionableComponents.some(function (earning) {
        return earning.code === "CAOHolidayAllowance";
      });

      // Add header
      result.push(["Allowance Name", "Amount", "Amount Type", "Interval", "Based On", "Individual Choice Budget", "Pensionable", "Conditions", "Description"]);

      // Process each remuneration record
      partnerData.forEach(function (record) {
        if (record.holidayAllowances && record.holidayAllowances.length > 0) {
          record.holidayAllowances.forEach(function (holidayAllowance) {
            var _holidayAllowance$amo;
            var amountDisplay = _this2.formatAmountDisplay(holidayAllowance.amount, holidayAllowance.amountType);
            result.push(["Holiday Allowance", amountDisplay, ((_holidayAllowance$amo = holidayAllowance.amountType) === null || _holidayAllowance$amo === void 0 ? void 0 : _holidayAllowance$amo.value) || "N/A", holidayAllowance.interval || "N/A", holidayAllowance.basedOn || "N/A", holidayAllowance.isPartOfIndividualChoiceBudget ? "true" : "", isHolidayAllowancePensionable ? "true" : "", "N/A", holidayAllowance.description || "N/A"]);
          });
        }
      });

      // If no holiday allowances found, add info message
      if (result.length === 1) {
        result.push(noDataFound);
      }
      return result;
    }

    /**
     * Generate allowances data for Excel format
     * Takes allowances.allowances objects from remuneration data
     * Returns formatted data with allowance name first, then component details in following columns
     */
  }, {
    key: "generateAllowances",
    value: function generateAllowances(partnerId) {
      var _this3 = this;
      var partnerData = this.getPartnerRemuneration(partnerId);
      var noDataFound = ["No allowances data available", "", "", "", "", "", "", "", ""];
      if (!partnerData || partnerData.length === 0) {
        return [noDataFound];
      }
      var result = [];

      // Check if allowances are pensionable for this partner
      var pensionableComponents = this.getPensionableComponents(partnerId);
      var isAllowancesPensionable = pensionableComponents.some(function (earning) {
        return earning.code === "CAOAllowances";
      });

      // Add main header with allowance name first, then component fields
      result.push(["Allowance Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"]);

      // Process each remuneration record
      partnerData.forEach(function (record) {
        if (record.allowances && record.allowances.hasAllowances && record.allowances.allowances) {
          // Process each allowance
          record.allowances.allowances.forEach(function (allowance) {
            var allowanceName = allowance.name || allowance.typeTitle || "Unnamed Allowance";

            // Process allowance components (lowest level)
            if (allowance.allowanceComponents && allowance.allowanceComponents.length > 0) {
              allowance.allowanceComponents.forEach(function (component) {
                var _component$amountType;
                var amountDisplay = _this3.formatAmountDisplay(component.amount, component.amountType);
                result.push([allowanceName, amountDisplay, ((_component$amountType = component.amountType) === null || _component$amountType === void 0 ? void 0 : _component$amountType.value) || "N/A", component.basedOn || "N/A", component.interval || "N/A", "",
                // Allowances are typically not part of ICB
                isAllowancesPensionable ? "true" : "", component.conditions || "N/A", component.description || "N/A"]);
              });
            } else {
              // No components, show allowance with no component data
              result.push([allowanceName, "N/A", "N/A", "N/A", "N/A", "", isAllowancesPensionable ? "true" : "", allowance.conditions || "N/A", "No components available"]);
            }
          });
        }
      });

      // If no allowances found, add info message
      if (result.length === 1) {
        result.push(noDataFound);
      }
      console.log(result);
      return result;
    }

    /**
     * Generate wage increments data for Excel format
     * Covers incrementData, retroactiveIncrementData, and periodicalData from wageIncrements
     */
  }, {
    key: "generateWageIncrementsData",
    value: function generateWageIncrementsData(partnerId) {
      var _this4 = this;
      return this.generateDataWithStandardStructure(partnerId, "No wage increments data available", ["Component Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.wageIncrements && record.wageIncrements.hasWageIncrements) {
            // Process regular increment data
            if (record.wageIncrements.incrementData && record.wageIncrements.incrementData.length > 0) {
              record.wageIncrements.incrementData.forEach(function (increment) {
                var amountDisplay = _this4.formatAmountDisplay(increment.percentage, {
                  value: increment.incrementType === "%" ? "Percentage" : "Amount",
                  otherValue: null
                });
                rows.push(["Regular Wage Increment", amountDisplay, increment.incrementType === "%" ? "Percentage" : "Amount", "Salary", "Effective ".concat(increment.date), "",
                // Wage increments typically not part of ICB
                "",
                // Pensionable - dummy column for consistency
                increment.appliesToAllFunctions ? "Applies to all functions" : "Function-specific", increment.explanation || "Wage increment of ".concat(increment.percentage).concat(increment.incrementType, " effective ").concat(increment.date)]);
              });
            }

            // Process retroactive increment data
            if (record.wageIncrements.retroactiveIncrementData && record.wageIncrements.retroactiveIncrementData.length > 0) {
              record.wageIncrements.retroactiveIncrementData.forEach(function (increment) {
                var amountDisplay = _this4.formatAmountDisplay(increment.percentage, {
                  value: increment.incrementType === "%" ? "Percentage" : "Amount",
                  otherValue: null
                });
                rows.push(["Retroactive Wage Increment", amountDisplay, increment.incrementType === "%" ? "Percentage" : "Amount", "Salary", "Effective ".concat(increment.date), "",
                // Wage increments typically not part of ICB
                "",
                // Pensionable - dummy column for consistency
                increment.appliesToAllFunctions ? "Applies to all functions" : "Function-specific", increment.explanation || "Retroactive wage increment of ".concat(increment.percentage).concat(increment.incrementType, " effective ").concat(increment.date)]);
              });
            }

            // Process periodical data
            if (record.wageIncrements.periodicalData && record.wageIncrements.periodicalData.length > 0) {
              record.wageIncrements.periodicalData.forEach(function (periodical) {
                rows.push(["Periodical Increment", "N/A", "Schedule-based", "Performance & Experience", "".concat(periodical.experiencePeriodAmount, " ").concat(periodical.experiencePeriodAmountType), "",
                // Wage increments typically not part of ICB
                "", // Pensionable - dummy column for consistency
                "".concat(periodical.eligibleType, " - ").concat(periodical.performanceMoment), periodical.periodicals || "Performance-based salary progression"]);
              });
            }
          }
        });
        return rows;
      });
    }

    /**
     * Retrieve ALL components that are pensionable; this list is used to mark the components
     */
  }, {
    key: "getPensionableComponents",
    value: function getPensionableComponents(partnerId) {
      var partnerData = this.getPartnerRemuneration(partnerId);
      var pensionData = partnerData.length > 0 ? partnerData[0].pensionData : null;

      //console.log('Pension Data:', pensionData);

      if (!pensionData || !pensionData.items || pensionData.items.length === 0) {
        return [];
      }
      var allPensionableEarnings = [];
      pensionData.items.forEach(function (item) {
        if (item.pensionableEarnings && item.pensionableEarnings.length > 0) {
          allPensionableEarnings.push.apply(allPensionableEarnings, _toConsumableArray(item.pensionableEarnings));
        }
      });

      // Enhance with component mapping
      var enhancedEarnings = this.enhancePensionableEarningsWithComponents(allPensionableEarnings);

      //console.log('Pensionable Earnings:', enhancedEarnings);
      return enhancedEarnings;
    }

    /**
     * Generate pension data for Excel format
     * Similar to holiday allowances - simple table format
     */
  }, {
    key: "generatePensionData",
    value: function generatePensionData(partnerId) {
      var _this5 = this;
      return this.generateDataWithStandardStructure(partnerId, "No pension data available", ["Component Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.pensionData && record.pensionData.hasPension && record.pensionData.items) {
            record.pensionData.items.forEach(function (item) {
              var _item$amountType;
              var amountDisplay = _this5.formatAmountDisplay(item.amount, item.amountType);

              // determine pensionable components
              var pensionableItems = item.pensionableEarnings && item.pensionableEarnings.length > 0 ? item.pensionableEarnings.map(function (pe) {
                return pe.title;
              }).join(", ") : "N/A";
              rows.push(["Pension", amountDisplay, ((_item$amountType = item.amountType) === null || _item$amountType === void 0 ? void 0 : _item$amountType.value) || "N/A", pensionableItems || "N/A", item.interval || "N/A", "",
              // Pension typically not part of ICB
              "",
              // Pensionable - dummy column for consistency
              "N/A", item.description || "N/A"]);
            });
          }
        });
        return rows;
      });
    }

    /**
     * Generate leave data for Excel format
     * Each subcomponent treated as a new row like allowances
     */
  }, {
    key: "generateLeaveData",
    value: function generateLeaveData(partnerId) {
      var _this6 = this;
      return this.generateDataWithStandardStructure(partnerId, "No leave data available", ["Leave Type", "Amount", "Amount Type", "Interval", "Based On", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];

        // Check if holiday/leave components are pensionable for this partner
        var pensionableComponents = _this6.getPensionableComponents(partnerId);
        var isLeavePensionable = pensionableComponents.some(function (earning) {
          return earning.code === "CAOHoliday";
        });

        // Helper function to process leave components
        var processLeaveComponents = function processLeaveComponents(components, typeName) {
          components.forEach(function (component) {
            var _component$amountType2;
            var amountDisplay = _this6.formatAmountDisplay(component.amount, component.amountType);
            rows.push([typeName, amountDisplay, ((_component$amountType2 = component.amountType) === null || _component$amountType2 === void 0 ? void 0 : _component$amountType2.value) || "N/A", component.interval || "N/A", "",
            // Based On - not in LeaveComponent interface
            component.isPartOfIndividualChoiceBudget ? "true" : "", isLeavePensionable ? "true" : "", component.conditions || "N/A", component.description || "N/A"]);
          });
        };
        partnerData.forEach(function (record) {
          if (record.leave) {
            var leave = record.leave;

            // Process each leave type using the helper function
            if (leave.hasWorkingHoursReduction && leave.workingHoursReductions) {
              processLeaveComponents(leave.workingHoursReductions, "Working Hours Reduction");
            }
            if (leave.hasVacationDays && leave.vacationDays) {
              processLeaveComponents(leave.vacationDays, "Vacation Days");
            }
            if (leave.hasSpecialLeave && leave.specialLeave) {
              processLeaveComponents(leave.specialLeave, "Special Leave");
            }
            if (leave.hasAdditionalWAZO && leave.additionalWAZO) {
              processLeaveComponents(leave.additionalWAZO, "Additional WAZO");
            }
            if (leave.hasMandatoryLeave && leave.mandatoryLeave) {
              processLeaveComponents(leave.mandatoryLeave, "Mandatory Leave");
            }
          }
        });
        return rows;
      });
    }

    /**
     * Generate sick pay data for Excel format
     * Each sick pay amount treated as a new row like allowances
     */
  }, {
    key: "generateSickPayData",
    value: function generateSickPayData(partnerId) {
      var _this7 = this;
      return this.generateDataWithStandardStructure(partnerId, "No sick pay data available", ["Component Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.sickPay && record.sickPay.sickPayAmounts && Array.isArray(record.sickPay.sickPayAmounts) && record.sickPay.sickPayAmounts.length > 0) {
            record.sickPay.sickPayAmounts.forEach(function (sickPayAmount) {
              var _sickPayAmount$sickPa, _sickPayAmount$amount;
              var amountDisplay = _this7.formatAmountDisplay(sickPayAmount.amount, sickPayAmount.amountType);
              rows.push(["Sick Pay (".concat(((_sickPayAmount$sickPa = sickPayAmount.sickPayYear) === null || _sickPayAmount$sickPa === void 0 ? void 0 : _sickPayAmount$sickPa.value) || "Unknown Year", ")"), amountDisplay, ((_sickPayAmount$amount = sickPayAmount.amountType) === null || _sickPayAmount$amount === void 0 ? void 0 : _sickPayAmount$amount.value) || "N/A", sickPayAmount.basedOn || "N/A", sickPayAmount.interval || "N/A", "",
              // Sick pay typically not part of ICB
              "",
              // Pensionable - dummy column for consistency
              "N/A", "N/A"]);
            });
          }
        });
        return rows;
      });
    }

    /**
     * Generate individual choice budget data for Excel format
     * Includes items and ICB column
     */
  }, {
    key: "generateIndividualChoiceBudget",
    value: function generateIndividualChoiceBudget(partnerId) {
      var _this8 = this;
      return this.generateDataWithStandardStructure(partnerId, "No individual choice budget data available", ["Component Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.individualChoiceBudget && record.individualChoiceBudget.hasIndividualChoiceBudget) {
            var _budget$amountType;
            var budget = record.individualChoiceBudget;

            // Add main budget entry
            var amountDisplay = _this8.formatAmountDisplay(budget.amount, budget.amountType);
            rows.push(["Individual Choice Budget", amountDisplay, ((_budget$amountType = budget.amountType) === null || _budget$amountType === void 0 ? void 0 : _budget$amountType.value) || "N/A", budget.basedOn || "N/A", budget.interval || "N/A", "true",
            // Main budget is always part of ICB
            "",
            // Pensionable - dummy column for consistency
            "N/A", "Main budget allocation"]);

            // Add items if they exist
            if (budget.items && budget.items.length > 0) {
              budget.items.forEach(function (item) {
                var _item$amountType2;
                var itemAmountDisplay = _this8.formatAmountDisplay(item.amount, item.amountType);
                var componentName = item.name === "Other" && item.typeOther ? item.typeOther : item.name || "Unnamed Component";
                rows.push([componentName, itemAmountDisplay, ((_item$amountType2 = item.amountType) === null || _item$amountType2 === void 0 ? void 0 : _item$amountType2.value) || "N/A", item.basedOn || "N/A", item.interval || "N/A", item.isPartOfIndividualChoiceBudget ? "true" : "", "",
                // Pensionable - dummy column for consistency
                "N/A",
                // No conditions field in the data structure
                item.description || "N/A"]);
              });
            }
          }
        });
        return rows;
      });
    }

    /**
     * Generate reimbursements data for Excel format
     * Each reimbursement treated as a new row like allowances
     */
  }, {
    key: "generateReimbursementsData",
    value: function generateReimbursementsData(partnerId) {
      var _this9 = this;
      return this.generateDataWithStandardStructure(partnerId, "No reimbursements data available", ["Reimbursement Name", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.reimbursementsData && record.reimbursementsData.hasReimbursements && record.reimbursementsData.reimbursements) {
            record.reimbursementsData.reimbursements.forEach(function (reimbursement) {
              var _reimbursement$amount;
              var amountDisplay = _this9.formatAmountDisplay(reimbursement.amount, reimbursement.amountType);
              rows.push([reimbursement.name || "Unnamed Reimbursement", amountDisplay, ((_reimbursement$amount = reimbursement.amountType) === null || _reimbursement$amount === void 0 ? void 0 : _reimbursement$amount.value) || "N/A", reimbursement.basedOn || "N/A", reimbursement.interval || "N/A", "",
              // Reimbursements typically not part of ICB
              "",
              // Pensionable - dummy column for consistency
              reimbursement.conditions || "N/A", reimbursement.description || "N/A"]);
            });
          }
        });
        return rows;
      });
    }

    /**
     * Generate payments data for Excel format
     * Each anniversary payment treated as a new row like allowances
     */
  }, {
    key: "generatePaymentsData",
    value: function generatePaymentsData(partnerId) {
      var _this0 = this;
      return this.generateDataWithStandardStructure(partnerId, "No payments data available", ["Payment Type", "Amount", "Amount Type", "Based On", "Interval", "Individual Choice Budget", "Pensionable", "Conditions", "Description"], function (partnerData) {
        var rows = [];
        partnerData.forEach(function (record) {
          if (record.paymentsData && record.paymentsData.hasAnniversaryPayments && record.paymentsData.anniversaryPayments) {
            record.paymentsData.anniversaryPayments.forEach(function (payment) {
              var _payment$amountType;
              var amountDisplay = _this0.formatAmountDisplay(payment.amount, payment.amountType);
              var paymentName = "Anniversary Payment (".concat(payment.amountOfYears, " years)");
              var basedOnText = payment.basedOn === "Other" && payment.basedOnOther ? payment.basedOnOther : payment.basedOn || "N/A";
              var conditionsText = payment.dateOrCondition === "Other" && payment.dateOrConditionOther ? payment.dateOrConditionOther : payment.dateOrCondition || "N/A";
              rows.push([paymentName, amountDisplay, ((_payment$amountType = payment.amountType) === null || _payment$amountType === void 0 ? void 0 : _payment$amountType.value) || "N/A", basedOnText, payment.interval || "N/A", "",
              // Anniversary payments typically not part of ICB
              "",
              // Pensionable - dummy column for consistency
              conditionsText, payment.description || "N/A"]);
            });
          }
        });
        return rows;
      });
    }

    /**
     * Extracts and concatenates all pensionable earnings from pension data items
     * @param pensionData The pension data object containing items with pensionable earnings
     * @returns Array of all pensionable earnings from all items with component mapping
     */
  }, {
    key: "extractAllPensionableEarnings",
    value: function extractAllPensionableEarnings(pensionData) {
      if (!pensionData || !pensionData.items || pensionData.items.length === 0) {
        return [];
      }
      var allPensionableEarnings = [];
      pensionData.items.forEach(function (item) {
        if (item.pensionableEarnings && item.pensionableEarnings.length > 0) {
          allPensionableEarnings.push.apply(allPensionableEarnings, _toConsumableArray(item.pensionableEarnings));
        }
      });

      // Enhance with component mapping
      return this.enhancePensionableEarningsWithComponents(allPensionableEarnings);
    }

    /**
     * Get all unique component types from pensionable earnings for a partner
     * @param partnerId Partner ID to filter data
     * @returns Array of unique component names
     */
  }, {
    key: "getPensionableComponentTypes",
    value: function getPensionableComponentTypes(partnerId) {
      var pensionableEarnings = this.getPensionableComponents(partnerId);
      var componentTypes = pensionableEarnings.map(function (earning) {
        return earning.component;
      }).filter(function (component, index, array) {
        return component && array.indexOf(component) === index;
      });
      return componentTypes;
    }

    /**
     * Check if a specific component type is pensionable for a partner
     * @param partnerId Partner ID to filter data
     * @param componentType The component type to check (e.g., "allowances", "holidayAllowance")
     * @returns True if the component type is pensionable
     */
  }, {
    key: "isComponentTypePensionable",
    value: function isComponentTypePensionable(partnerId, componentType) {
      var componentTypes = this.getPensionableComponentTypes(partnerId);
      return componentTypes.includes(componentType);
    }
  }]);
}();

// Create a singleton instance
var dataLoader = new DataLoaderService();

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "7e68e84774d219c9c3b1.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   loadPartners: function() { return /* binding */ loadPartners; },
/* harmony export */   loadRemuneration: function() { return /* binding */ loadRemuneration; },
/* harmony export */   onPartnerSelect: function() { return /* binding */ onPartnerSelect; }
/* harmony export */ });
/* harmony import */ var _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../services/dataLoaderService */ "./src/services/dataLoaderService.ts");
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */



// Global variables
var partners = [];
var selectedPartnerId = "";

// The initialize function must be run each time a new page is loaded
Office.onReady(function () {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  // Wire up event handlers
  document.getElementById("load-partners").onclick = loadPartners;
  document.getElementById("partner-select").onchange = onPartnerSelect;
  document.getElementById("load-remuneration").onclick = loadRemuneration;

  // Load partners on startup
  loadPartners();
});

/**
 * Load partners and populate the dropdown
 */
function loadPartners() {
  return _loadPartners.apply(this, arguments);
}

/**
 * Handle partner selection
 */
function _loadPartners() {
  _loadPartners = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var select, _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          console.log('Starting to load partners...');
          updateStatus("Loading partners from JSON data...", "loading");
          _context.p = 1;
          _context.n = 2;
          return _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.loadData();
        case 2:
          // Get partners from the loaded data
          partners = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.getPartners();
          console.log('Partners loaded:', partners);
          select = document.getElementById("partner-select");
          if (select) {
            // Clear existing options
            select.innerHTML = '<option value="">Select a partner...</option>';

            // Add partner options
            partners.forEach(function (partner) {
              console.log('Adding partner to dropdown:', partner);
              var option = document.createElement('option');
              option.value = partner.id;
              option.textContent = partner.businessName;
              select.appendChild(option);
            });

            // Enable the dropdown
            select.disabled = false;
          }
          updateStatus("Loaded ".concat(partners.length, " partners"), "success");
          _context.n = 4;
          break;
        case 3:
          _context.p = 3;
          _t = _context.v;
          console.error("Error loading partners:", _t);
          updateStatus("Error loading partners: ".concat(_t.message), "error");
        case 4:
          return _context.a(2);
      }
    }, _callee, null, [[1, 3]]);
  }));
  return _loadPartners.apply(this, arguments);
}
function onPartnerSelect() {
  var select = document.getElementById("partner-select");
  if (select) {
    selectedPartnerId = select.value;
    var loadButton = document.getElementById("load-remuneration");
    if (loadButton) {
      if (selectedPartnerId) {
        loadButton.classList.remove('disabled');
      } else {
        loadButton.classList.add('disabled');
      }
    }
    if (selectedPartnerId) {
      var selectedPartner = partners.find(function (p) {
        return p.id === selectedPartnerId;
      });
      updateStatus("Selected partner: ".concat((selectedPartner === null || selectedPartner === void 0 ? void 0 : selectedPartner.businessName) || selectedPartnerId), "success");
    } else {
      updateStatus("Select a partner to continue", "loading");
    }
  }
}

/**
 * Load remuneration data for the selected partner
 */
function loadRemuneration() {
  return _loadRemuneration.apply(this, arguments);
}

/**
 * Load data into the ILB sheet with proper structure
 */
function _loadRemuneration() {
  _loadRemuneration = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var loadButton, selectedPartner, remunerationData, _t2;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          loadButton = document.getElementById("load-remuneration");
          if (!(loadButton && loadButton.classList.contains('disabled'))) {
            _context3.n = 1;
            break;
          }
          return _context3.a(2);
        case 1:
          if (selectedPartnerId) {
            _context3.n = 2;
            break;
          }
          updateStatus("Please select a partner first", "error");
          return _context3.a(2);
        case 2:
          selectedPartner = partners.find(function (p) {
            return p.id === selectedPartnerId;
          });
          updateStatus("Loading remuneration and allowances for ".concat(selectedPartner === null || selectedPartner === void 0 ? void 0 : selectedPartner.businessName, "..."), "loading");
          _context3.p = 3;
          // Get remuneration data for the partner
          remunerationData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.getPartnerRemuneration(selectedPartnerId);
          if (!(!remunerationData || remunerationData.length === 0)) {
            _context3.n = 4;
            break;
          }
          updateStatus("No remuneration available for this partner", "error");
          return _context3.a(2);
        case 4:
          _context3.n = 5;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(context) {
              return _regenerator().w(function (_context2) {
                while (1) switch (_context2.n) {
                  case 0:
                    _context2.n = 1;
                    return loadDataToILBSheet(context, selectedPartnerId, (selectedPartner === null || selectedPartner === void 0 ? void 0 : selectedPartner.businessName) || 'Unknown');
                  case 1:
                    updateStatus("Remuneration and allowances loaded successfully for ".concat(selectedPartner === null || selectedPartner === void 0 ? void 0 : selectedPartner.businessName, " (").concat(remunerationData.length, " records)"), "success");
                  case 2:
                    return _context2.a(2);
                }
              }, _callee2);
            }));
            return function (_x9) {
              return _ref.apply(this, arguments);
            };
          }());
        case 5:
          _context3.n = 7;
          break;
        case 6:
          _context3.p = 6;
          _t2 = _context3.v;
          console.error("Error loading remuneration:", _t2);
          updateStatus("Error loading remuneration: ".concat(_t2.message), "error");
        case 7:
          return _context3.a(2);
      }
    }, _callee3, null, [[3, 6]]);
  }));
  return _loadRemuneration.apply(this, arguments);
}
function loadDataToILBSheet(_x, _x2, _x3) {
  return _loadDataToILBSheet.apply(this, arguments);
}
/**
 * Clears all existing named ranges in the workbook
 * @param context - The Excel request context
 */
function _loadDataToILBSheet() {
  _loadDataToILBSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(context, partnerId, partnerName) {
    var worksheet, worksheets, existingWorksheet, usedRange, metadata, metadataRange, workingHoursCell, namedItem, currentDataRow, pensionableComponents, holidayAllowancesData, leaveData, individualChoiceBudgetData, pensionData, sickPayData, allowancesData, reimbursementsData, paymentsData, columnB, cellB8, columnCF;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          console.log("Preparing to load data into ILB sheet for partner ID: ".concat(partnerId));

          // Get or create ILB worksheet using a more reliable approach
          // Load all worksheets to check if ILB exists
          worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          _context4.n = 1;
          return context.sync();
        case 1:
          // Check if ILB worksheet already exists
          existingWorksheet = worksheets.items.find(function (ws) {
            return ws.name === "ILB";
          });
          if (!existingWorksheet) {
            _context4.n = 2;
            break;
          }
          worksheet = existingWorksheet;
          //console.log(`Found existing worksheet: ${worksheet.name}`);
          _context4.n = 3;
          break;
        case 2:
          console.log("ILB worksheet not found, creating a new one.");
          worksheet = worksheets.add("ILB");
          //console.log(`Created new worksheet: ${worksheet.name}`);
          _context4.n = 3;
          return context.sync();
        case 3:
          console.log("Clearing existing data in ILB sheet...");

          // Clear existing named ranges first
          _context4.n = 4;
          return clearExistingNamedRanges(context);
        case 4:
          // Clear the entire sheet
          usedRange = worksheet.getUsedRange();
          usedRange.load("address");
          _context4.n = 5;
          return context.sync();
        case 5:
          if (!(usedRange.address !== "A1:A1" || usedRange.values[0][0] !== "")) {
            _context4.n = 6;
            break;
          }
          usedRange.clear();
          _context4.n = 6;
          return context.sync();
        case 6:
          console.log("Loading data for partner: ".concat(partnerName, " into ILB sheet"));

          // Generate metadata (rows 1-10)
          metadata = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateMetadata(partnerId);
          console.log('Metadata generated:', metadata);
          metadataRange = worksheet.getRange("A1:C10"); // Dynamically set range for metadata
          metadataRange.values = metadata;

          // Format metadata section
          metadataRange.format.font.bold = true;
          metadataRange.getColumn(0).format.fill.color = "#E3F2FD";
          setAllBorders(metadataRange, false);

          // Set named ranges using the workbook's named items collection
          workingHoursCell = worksheet.getRange("B7");
          namedItem = context.workbook.names.add("WorkingHoursPerWeek", workingHoursCell);
          _context4.n = 7;
          return context.sync();
        case 7:
          currentDataRow = 11; // A12 starts at row index 9 (0-based)
          // Load components that are pensionable
          pensionableComponents = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.getPensionableComponents(partnerId);
          console.log('Pensionable Components:', pensionableComponents);

          // Generate holiday allowances data (starting from A10)  
          holidayAllowancesData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateHolidayAllowances(partnerId);
          console.log('Holiday Allowances Data:', holidayAllowancesData);
          if (!(holidayAllowancesData.length > 0)) {
            _context4.n = 9;
            break;
          }
          _context4.n = 8;
          return addComponent(worksheet, "Holiday Allowances", holidayAllowancesData, currentDataRow);
        case 8:
          currentDataRow = _context4.v;
        case 9:
          // Generate leave data
          leaveData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateLeaveData(partnerId);
          if (!(leaveData.length > 0)) {
            _context4.n = 11;
            break;
          }
          _context4.n = 10;
          return addComponent(worksheet, "Leave", leaveData, currentDataRow);
        case 10:
          currentDataRow = _context4.v;
        case 11:
          // Generate individual choice budget data
          individualChoiceBudgetData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateIndividualChoiceBudget(partnerId);
          if (!(individualChoiceBudgetData.length > 0)) {
            _context4.n = 13;
            break;
          }
          _context4.n = 12;
          return addComponent(worksheet, "Individual Choice Budget", individualChoiceBudgetData, currentDataRow);
        case 12:
          currentDataRow = _context4.v;
        case 13:
          // Generate pension data
          pensionData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generatePensionData(partnerId);
          if (!(pensionData.length > 0)) {
            _context4.n = 15;
            break;
          }
          _context4.n = 14;
          return addComponent(worksheet, "Pension", pensionData, currentDataRow);
        case 14:
          currentDataRow = _context4.v;
        case 15:
          // Generate sick pay data
          sickPayData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateSickPayData(partnerId);
          if (!(sickPayData.length > 0)) {
            _context4.n = 17;
            break;
          }
          _context4.n = 16;
          return addComponent(worksheet, "Sick Pay", sickPayData, currentDataRow);
        case 16:
          currentDataRow = _context4.v;
        case 17:
          // Generate allowances data
          allowancesData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateAllowances(partnerId);
          if (!(allowancesData.length > 0)) {
            _context4.n = 19;
            break;
          }
          _context4.n = 18;
          return addComponent(worksheet, "Allowances", allowancesData, currentDataRow);
        case 18:
          currentDataRow = _context4.v;
        case 19:
          // Generate reimbursements data
          reimbursementsData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generateReimbursementsData(partnerId);
          if (!(reimbursementsData.length > 0)) {
            _context4.n = 21;
            break;
          }
          _context4.n = 20;
          return addComponent(worksheet, "Reimbursements", reimbursementsData, currentDataRow);
        case 20:
          currentDataRow = _context4.v;
        case 21:
          // Generate payments data (anniversary payments)
          paymentsData = _services_dataLoaderService__WEBPACK_IMPORTED_MODULE_0__.dataLoader.generatePaymentsData(partnerId);
          if (!(paymentsData.length > 0)) {
            _context4.n = 23;
            break;
          }
          _context4.n = 22;
          return addComponent(worksheet, "Payments", paymentsData, currentDataRow);
        case 22:
          currentDataRow = _context4.v;
        case 23:
          // Auto-fit columns
          worksheet.getUsedRange().format.autofitColumns();

          // Set specific column widths and word wrap
          columnB = worksheet.getRange("B:B");
          columnB.format.columnWidth = 100;
          cellB8 = worksheet.getCell(7, 1); // B8 (0-based indexing: row 7, col 1)
          cellB8.format.wrapText = false;
          columnCF = worksheet.getRange("C:G");
          columnCF.format.columnWidth = 100;
          _context4.n = 24;
          return context.sync();
        case 24:
          console.log("Data loaded successfully into ILB sheet for ".concat(partnerName));
        case 25:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _loadDataToILBSheet.apply(this, arguments);
}
function clearExistingNamedRanges(_x4) {
  return _clearExistingNamedRanges.apply(this, arguments);
}
/**
 * Sets all borders (edge and inside) for a given range
 * @param range - The Excel range to apply borders to
 * @param includeInside - Whether to include inside borders (default: true)
 */
function _clearExistingNamedRanges() {
  _clearExistingNamedRanges = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(context) {
    var namedItems, itemsToDelete, _iterator, _step, item;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.n) {
        case 0:
          console.log("Clearing existing named ranges...");
          namedItems = context.workbook.names;
          namedItems.load("items");
          _context5.n = 1;
          return context.sync();
        case 1:
          // Remove all existing named ranges
          itemsToDelete = namedItems.items.slice(); // Create a copy to avoid modification during iteration
          _iterator = _createForOfIteratorHelper(itemsToDelete);
          try {
            for (_iterator.s(); !(_step = _iterator.n()).done;) {
              item = _step.value;
              console.log("Deleting named range: ".concat(item.name));
              item.delete();
            }
          } catch (err) {
            _iterator.e(err);
          } finally {
            _iterator.f();
          }
          if (!(itemsToDelete.length > 0)) {
            _context5.n = 3;
            break;
          }
          _context5.n = 2;
          return context.sync();
        case 2:
          console.log("Cleared ".concat(itemsToDelete.length, " existing named ranges"));
          _context5.n = 4;
          break;
        case 3:
          console.log("No existing named ranges to clear");
        case 4:
          return _context5.a(2);
      }
    }, _callee5);
  }));
  return _clearExistingNamedRanges.apply(this, arguments);
}
function setAllBorders(range) {
  var includeInside = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : true;
  range.format.borders.getItem("EdgeTop").style = "Continuous";
  range.format.borders.getItem("EdgeBottom").style = "Continuous";
  range.format.borders.getItem("EdgeLeft").style = "Continuous";
  range.format.borders.getItem("EdgeRight").style = "Continuous";
  if (includeInside) {
    range.format.borders.getItem("InsideVertical").style = "Continuous";
    range.format.borders.getItem("InsideHorizontal").style = "Continuous";
  }
}

/**
 * Generic function to add an ILB component table to Excel with consistent formatting
 * @param worksheet - The Excel worksheet
 * @param title - The component title (e.g., "Allowances", "Holiday Allowances")
 * @param data - The data array with headers as first row
 * @param startRow - The row index to start placing the component (0-based)
 * @returns The next available row index after this component
 */
function addComponent(_x5, _x6, _x7, _x8) {
  return _addComponent.apply(this, arguments);
}
/**
 * Updates the status message
 */
function _addComponent() {
  _addComponent = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(worksheet, title, data, startRow) {
    var currentRow, titleRange, dataEndCol, dataRange, headerRange, conditionsCol, conditionsRange, descriptionCol, descriptionRange, col, colRange;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.n) {
        case 0:
          if (!(!data || data.length === 0)) {
            _context6.n = 1;
            break;
          }
          return _context6.a(2, startRow);
        case 1:
          currentRow = startRow; // Add component title - only 1 column
          titleRange = worksheet.getRangeByIndexes(currentRow, 0, 1, 1);
          titleRange.values = [[title]];
          titleRange.format.font.bold = true;
          titleRange.format.fill.color = "#ADD8E6"; // Light blue
          setAllBorders(titleRange, false);
          currentRow++;

          // Add data (including headers)
          dataEndCol = Math.max(0, Math.max.apply(Math, _toConsumableArray(data.map(function (row) {
            return row.length;
          }))) - 1);
          dataRange = worksheet.getRangeByIndexes(currentRow, 0, data.length, dataEndCol + 1);
          dataRange.values = data;

          // Format all data cells with borders
          setAllBorders(dataRange);

          // Format header row (first row of data) as bold
          if (data.length > 0) {
            headerRange = worksheet.getRangeByIndexes(currentRow, 0, 1, dataEndCol + 1);
            headerRange.format.font.bold = true;
          }

          // Special formatting for description and conditions columns (last two columns if they exist)
          if (dataEndCol >= 6) {
            // Only if we have at least 7 columns (including description and conditions)
            // Format conditions column (2nd to last)
            conditionsCol = dataEndCol - 1;
            conditionsRange = worksheet.getRangeByIndexes(currentRow, conditionsCol, data.length, 1);
            conditionsRange.format.wrapText = true;
            conditionsRange.format.columnWidth = 400;
            conditionsRange.format.verticalAlignment = "Top";

            // Format description column (last column) 
            descriptionCol = dataEndCol;
            descriptionRange = worksheet.getRangeByIndexes(currentRow, descriptionCol, data.length, 1);
            descriptionRange.format.wrapText = true;
            descriptionRange.format.columnWidth = 400;
            descriptionRange.format.verticalAlignment = "Top";
          }

          // Auto-fit other columns
          for (col = 0; col <= Math.min(dataEndCol - 2, dataEndCol); col++) {
            colRange = worksheet.getRangeByIndexes(currentRow, col, data.length, 1);
            colRange.format.autofitColumns();
            colRange.format.verticalAlignment = "Top";
          }
          return _context6.a(2, currentRow + data.length + 1);
      }
    }, _callee6);
  }));
  return _addComponent.apply(this, arguments);
}
function updateStatus(message, type) {
  var statusDiv = document.getElementById("status");
  if (statusDiv) {
    statusDiv.textContent = message;
    statusDiv.className = "status ".concat(type);
  }
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\n\n<!DOCTYPE html>\n<html lang=\"en\">\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>ILB Add-in</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">        \n        <h1 class=\"ms-font-xl\">ILB tool</h1>\n    </header>\n    <section id=\"sideload-msg\" class=\"ms-bgColor-neutralLighter ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" rel=\"noopener noreferrer\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\n    </section>\n    <main id=\"app-body\" class=\"ms-bgColor-neutralLighter ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Inlenersbeloning Data Loader</h2>\n        \n        <!-- Partner Selection Section -->\n        <div class=\"api-section\">\n            <h3 class=\"ms-font-l\">Select Partner</h3>\n            \n            <div class=\"input-group\">\n                <label for=\"partner-select\" class=\"ms-font-m\">Partner:</label>\n                <select id=\"partner-select\" class=\"ms-TextField-field\" disabled>\n                    <option value=\"\">Loading partners...</option>\n                </select>\n            </div>\n            \n            <div class=\"button-group\">\n                <div role=\"button\" id=\"load-partners\" class=\"ms-Button ms-Button--secondary ms-font-m ms-fontWeight-semibold\">\n                    <span class=\"ms-Button-label ms-fontAlign-center\">Refresh Partners</span>\n                </div>\n                \n                <div role=\"button\" id=\"load-remuneration\" class=\"ms-Button ms-Button--primary ms-font-m ms-fontWeight-semibold disabled\">\n                    <span class=\"ms-Button-label ms-fontAlign-center\">Load Remuneration Data</span>\n                </div>\n            </div>\n        </div>\n\n        <!-- Status Section -->\n        <div id=\"status\" class=\"status\">Loading partners from JSON data...</div>\n    </main>\n        <p><label id=\"item-subject\"></label></p>\n    </main>\n</body>\n\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map
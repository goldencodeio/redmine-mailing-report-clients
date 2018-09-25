function formatDate(date) {
  return date.toJSON().split('T').shift();
}

function getDateRage(startDate, finalDate) {
  return '><' + formatDate(startDate) + '|' + formatDate(finalDate);
}

function stringRating(number) {
  if (number < 3) return 'плохо';
  else if (number < 4) return 'удовлетворительно';
  else if (number < 4.5) return 'хорошо';
  else if (number < 5) return 'очень хорошо';
  else if (number >= 5) return 'отлично';
}

function stringMonth(date) {
  var month = date.getMonth();
  var rusMonth = '';
  switch (month) {
    case 0: rusMonth = 'январь'; break;
    case 1: rusMonth = 'февраль'; break;
    case 2: rusMonth = 'март'; break;
    case 3: rusMonth = 'апрель'; break;
    case 4: rusMonth = 'май'; break;
    case 5: rusMonth = 'июнь'; break;
    case 6: rusMonth = 'июль'; break;
    case 7: rusMonth = 'август'; break;
    case 8: rusMonth = 'сентябрь'; break;
    case 9: rusMonth = 'октябрь'; break;
    case 10: rusMonth = 'ноябрь'; break;
    case 11: rusMonth = 'декабрь'; break;
    default: rusMonth = '';
  }
  return rusMonth;
}

// function getDateRangeWithTime(startDate, finalDate) {
//   var startDate = startDate.toJSON().split('.').shift() + 'Z';
//   var finalDate = finalDate.toJSON().split('.').shift() + 'Z';
//   return '><' + startDate + '|' + finalDate;
// }
//
// function getHoursByRange(startDate, finalDate) {
//   return (finalDate.getTime() - startDate.getTime()) / (1000 * 60 * 60) - 1;
// }

function filterUniqueArray(arr) {
  return arr.sort(function(a,b){return a.id - b.id;}).reduce(function(arr, el){
    if(!arr.length || arr[arr.length - 1].id !== el.id) arr.push(el);
    return arr;
  }, []);
}

if (!Array.prototype.find) {
  Array.prototype.find = function(predicate) {
    if (this == null) {
      throw new TypeError('Array.prototype.find called on null or undefined');
    }
    if (typeof predicate !== 'function') {
      throw new TypeError('predicate must be a function');
    }
    var list = Object(this);
    var length = list.length >>> 0;
    var thisArg = arguments[1];
    var value;

    for (var i = 0; i < length; i++) {
      value = list[i];
      if (predicate.call(thisArg, value, i, list)) {
        return value;
      }
    }
    return undefined;
  };
}

if (!Array.prototype.findIndex) {
  Array.prototype.findIndex = function(predicate) {
    if (this == null) {
      throw new TypeError('Array.prototype.findIndex called on null or undefined');
    }
    if (typeof predicate !== 'function') {
      throw new TypeError('predicate must be a function');
    }
    var list = Object(this);
    var length = list.length >>> 0;
    var thisArg = arguments[1];
    var value;

    for (var i = 0; i < length; i++) {
      value = list[i];
      if (predicate.call(thisArg, value, i, list)) {
        return i;
      }
    }
    return -1;
  };
}

if (!Object.assign) {
  Object.defineProperty(Object, 'assign', {
    enumerable: false,
    configurable: true,
    writable: true,
    value: function(target, firstSource) {
      'use strict';
      if (target === undefined || target === null) {
        throw new TypeError('Cannot convert first argument to object');
      }

      var to = Object(target);
      for (var i = 1; i < arguments.length; i++) {
        var nextSource = arguments[i];
        if (nextSource === undefined || nextSource === null) {
          continue;
        }

        var keysArray = Object.keys(Object(nextSource));
        for (var nextIndex = 0, len = keysArray.length; nextIndex < len; nextIndex++) {
          var nextKey = keysArray[nextIndex];
          var desc = Object.getOwnPropertyDescriptor(nextSource, nextKey);
          if (desc !== undefined && desc.enumerable) {
            to[nextKey] = nextSource[nextKey];
          }
        }
      }
      return to;
    }
  });
}

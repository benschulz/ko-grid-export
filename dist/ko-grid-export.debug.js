/**
 * @license Copyright (c) 2015, Ben Schulz
 * License: BSD 3-clause (http://opensource.org/licenses/BSD-3-Clause)
 */
;(function(factory) {
    if (typeof define === 'function' && define['amd'])
        define(['ko-grid', 'ko-data-source', 'ko-indexed-repeat', 'knockout'], factory);
    else
        window['ko-grid-export'] = factory(window.ko.bindingHandlers['grid']);
} (function(ko_grid) {
var ko_grid_export_export, ko_grid_export;

var toolbar = 'ko-grid-toolbar';
ko_grid_export_export = function (module, koGrid) {
  var extensionId = 'ko-grid-export'.indexOf('/') < 0 ? 'ko-grid-export' : 'ko-grid-export'.substring(0, 'ko-grid-export'.indexOf('/'));
  var document = window.document, html = document.documentElement;
  var Blob = window.Blob, URL = window.URL;
  koGrid.defineExtension(extensionId, {
    dependencies: [toolbar],
    initializer: function (template) {
      template.into('left-toolbar').insert(' <button class="ko-grid-toolbar-button ko-grid-excel-export">Excel Export</button> ');
    },
    Constructor: function ExportExtension(bindingValue, config, grid) {
      grid.rootElement.addEventListener('click', function (e) {
        if (e.target.classList.contains('ko-grid-excel-export'))
          supplyExcelExport(grid);
      });
    }
  });
  return koGrid.declareExtensionAlias('export', extensionId);
  function supplyExcelExport(grid) {
    var columns = grid.columns.displayed().filter(function (c) {
      return !!c.property;
    });
    var valueSelector = grid.data.valueSelector;
    grid.data.source.streamValues(function (q) {
      return q.filteredBy(grid.data.predicate).sortedBy(grid.data.comparator);
    }).then(function (values) {
      return values.map(function (value) {
        return '<tr>' + columns.map(function (c) {
          return '<td>' + valueSelector(value[c.property]) + '</td>';
        }).join('') + '</tr>';
      }).reduce(function (a, b) {
        return a + b;
      }, '');
    }).then(function (data) {
      var excelDocument = [
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
        '  <head></head>' + '  <body>',
        '    <table>',
        '      <thead>',
        '        <tr>' + columns.map(function (c) {
          return '<th>' + c.label() + '</th>';
        }).join('') + '</tr>',
        // TODO groups
        '      </thead>',
        '      <tbody>',
        data,
        '      </tbody>',
        '    </table>',
        '  </body>',
        '</html>'
      ].join('');
      var saveOrOpen = window.navigator.msSaveOrOpenBlob ? window.navigator.msSaveOrOpenBlob.bind(window.navigator) : function (blob, name) {
        var url = URL.createObjectURL(blob);
        var anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = name;
        var event = document.createEvent('MouseEvents');
        event.initEvent('click', true, true);
        anchor.dispatchEvent(event);
        // TODO there needs to be a cleaner way to do this
        window.setTimeout(function () {
          return URL.revokeObjectURL(url);
        });
      };
      saveOrOpen(new Blob([excelDocument], { type: 'application/vnd.ms-excel;charset=utf-8' }), 'excel-export.xls');
    });
  }
}({}, ko_grid);
ko_grid_export = function (main) {
  return main;
}(ko_grid_export_export);return ko_grid_export;
}));
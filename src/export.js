'use strict';

var toolbar = 'ko-grid-toolbar';

define(['module', 'ko-grid', toolbar], function (module, koGrid) {
    var extensionId = module.id.indexOf('/') < 0 ? module.id : module.id.substring(0, module.id.indexOf('/'));

    var document = window.document,
        html = document.documentElement;

    var Blob = window.Blob,
        URL = window.URL;

    koGrid.defineExtension(extensionId, {
        dependencies: [toolbar],
        initializer: template => {
            template.into('left-toolbar').insert(' <button class="ko-grid-toolbar-button ko-grid-excel-export">Excel Export</button> ');
        },
        Constructor: function ExportExtension(bindingValue, config, grid) {
            grid.rootElement.addEventListener('click', e => {
                if (e.target.classList.contains('ko-grid-excel-export'))
                    supplyExcelExport(grid);
            });
        }
    });

    return koGrid.declareExtensionAlias('export', extensionId);

    function supplyExcelExport(grid) {
        var columns = grid.columns.displayed().filter(c => !!c.property);

        var valueSelector = grid.data.valueSelector;

        grid.data.source.streamValues(q => q.filteredBy(grid.data.predicate).sortedBy(grid.data.comparator))
            .then(values => values
                .map(value=> '<tr>' + columns.map(c => '<td>' + valueSelector(value[c.property]) + '</td>').join('') + '</tr>')
                .reduce((a, b) => a + b, ''))
            .then(data => {
                var excelDocument = [
                    '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
                    '  <head></head>' +
                    '  <body>',
                    '    <table>',
                    '      <thead>',
                    '        <tr>' + columns.map(c => '<th>' + c.label() + '</th>').join('') + '</tr>', // TODO groups
                    '      </thead>',
                    '      <tbody>',
                    data,
                    '      </tbody>',
                    '    </table>',
                    '  </body>',
                    '</html>'
                ].join('');

                var saveOrOpen = window.navigator.msSaveOrOpenBlob ? window.navigator.msSaveOrOpenBlob.bind(window.navigator) :
                    function (blob, name) {
                        var url = URL.createObjectURL(blob);
                        var anchor = document.createElement('a');
                        anchor.href = url;
                        anchor.download = name;
                        var event = document.createEvent('MouseEvents');
                        event.initEvent('click', true, true);
                        anchor.dispatchEvent(event);
                        // TODO there needs to be a cleaner way to do this
                        window.setTimeout(() => URL.revokeObjectURL(url));
                    };

                saveOrOpen(
                    new Blob([excelDocument], {type: 'application/vnd.ms-excel;charset=utf-8'}),
                    'excel-export.xls');
            });
    }
});

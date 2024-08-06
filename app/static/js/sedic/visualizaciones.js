function tipoFormatter(value, row, index) {
    var tipos = {{ tipos | tojson
}};
tipo = tipos.find(e => e.nombre === row.tipo)['valor']
return [tipo].join('')
  }

function dateFormatter(value, row, index) {
    fecha = JSON.stringify(row.fecha).replaceAll('"', '').substring(0, 10);
    return [fecha].join('')
}

function actionsFormatter(value, row, index) {
    let referencia = JSON.stringify(row.referencia).replaceAll('"', '');
    let pos = referencia.indexOf(':');
    let url = '';
    url = referencia.substring(pos + 1);

    return [
        '<a class="edit" href="#modalEdit" title="Editar" data-bs-toggle="modal">',
        '<i class="bi bi-pencil"></i>',
        '</a>&nbsp;',
        '<a class="preview" href="' + url + '" target="_blank" title="Vista Previa">',
        '<i class="bi bi-eye"></i>',
        '</a>&nbsp;',
        '<a class="delete" href="#modalDelete" title="Eliminar" data-bs-toggle="modal">',
        '<i class="bi bi-trash"></i>',
        '</a>'
    ].join('')
}

window.actionsEvents = {
    'click .edit': function (e, value, row, index) {
        let referencia = JSON.stringify(row.referencia).replaceAll('"', '');
        let pos = referencia.indexOf(':');
        document.getElementById('referencia').value = referencia.substr(pos + 1);
        if (row.tipo.substring(2, 3) == 'M')
            document.getElementById('metabase').value = referencia.substring(0, pos);
        else
            document.getElementById('metabase').value = 1;
        document.getElementById('formEdit').action = "{{url_for('Admin.Visualizaciones.action')}}" + '?dbAction=update';
        document.getElementById('titleEdit').innerHTML = '<i class="bi bi-pencil"></i>Editar Visualización';
        document.getElementById('codigo0').value = JSON.stringify(row.codigo).replaceAll('"', '');
        document.getElementById('codigo').value = JSON.stringify(row.codigo).replaceAll('"', '');
        document.getElementById('tipo').value = JSON.stringify(row.tipo).replaceAll('"', '');
        document.getElementById('nombre_corto').value = JSON.stringify(row.nombre_corto).replaceAll('"', '');
        document.getElementById('nombre').value = JSON.stringify(row.nombre).replaceAll('"', '');
        document.getElementById('descripcion').value = JSON.stringify(row.descripcion).replaceAll('"', '');
        //document.getElementById('referencia').value = JSON.stringify(row.referencia).replaceAll('"', '');
        document.getElementById('autor').value = JSON.stringify(row.autor).replaceAll('"', '');
        document.getElementById('encargado').value = JSON.stringify(row.encargado).replaceAll('"', '');
        document.getElementById('agrocadena').value = JSON.stringify(row.agrocadena).replaceAll('"', '');
        let et = JSON.stringify(row.etiquetas).replaceAll('"', '');
        $('#etiquetas').tagsinput('removeAll');
        $('#etiquetas').tagsinput('add', et.substr(1, et.length - 2));
        $('#etiquetas').tagsinput('refresh');
    },
    'click .delete': function (e, value, row, index) {
        const msg = 'Se eliminará la visualización:<br><i>' + JSON.stringify(row.nombre).replaceAll('"', '') + '</i>';
        const url = "{{url_for('Admin.Visualizaciones.delete')}}" + '?codigo=' + JSON.stringify(row.codigo).replaceAll('"', '');
        document.getElementById('titleDelete').innerHTML = '<em class="bi bi-trash"></em>Eliminar Visualización';
        document.getElementById('msgDelete').innerHTML = msg;
        document.getElementById('formDelete').action = url;
    }
}

$(function () {
    // New record
    $('#new').click(function () {
        document.getElementById('formEdit').action = "{{url_for('Admin.Visualizaciones.action')}}" + '?dbAction=create';
        document.getElementById('titleEdit').innerHTML = '<i class="bi bi-plus-lg"></i>Crear Visualización';
        document.getElementById('codigo').value = '';
        document.getElementById('tipo').value = '';
        document.getElementById('nombre_corto').value = '';
        document.getElementById('nombre').value = '';
        document.getElementById('descripcion').value = '';
        document.getElementById('referencia').value = '';
        document.getElementById('autor').value = '';
        document.getElementById('encargado').value = '';
        document.getElementById('agrocadena').value = '';
        $('#etiquetas').tagsinput('removeAll');
    });

    $('#table').bootstrapTable('destroy').bootstrapTable({
        exportDataType: 'all',
        exportTypes: ['json', 'xml', 'csv', 'txt', 'sql', 'excel'],
        locale: 'es-ES',
        pageSize: 10,
        data: {{ datos| tojson }}
    });
  })

$(document).ready(function () {
    // Inicializa DataTable
    var table = $('#inventoryTable').DataTable({
        "order": [[0, "desc"]],
        "pageLength": 25,
        "lengthChange": false
    });

    // Poblar filtros dinámicamente
    function poblarFiltros() {
        $('#inventoryTable tbody tr').each(function () {
            var tecnico = $(this).data('tecnico');
            var oficina = $(this).data('oficina');
            var marca = $(this).data('marca');

            if (tecnico && $('#filterTecnico option[value="' + tecnico + '"]').length === 0) {
                $('#filterTecnico').append('<option value="' + tecnico + '">' + tecnico + '</option>');
            }
            if (oficina && $('#filterOficina option[value="' + oficina + '"]').length === 0) {
                $('#filterOficina').append('<option value="' + oficina + '">' + oficina + '</option>');
            }
            if (marca && $('#filterMarca option[value="' + marca + '"]').length === 0) {
                $('#filterMarca').append('<option value="' + marca + '">' + marca + '</option>');
            }
        });
    }
    poblarFiltros();

    // Función de filtro
    $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
        var row = $('#inventoryTable tbody tr').eq(dataIndex);
        var tecnico = $('#filterTecnico').val();
        var oficina = $('#filterOficina').val();
        var marca = $('#filterMarca').val();
        var duplicado = $('#filterDuplicados').val();
        var searchText = $('#searchText').val().toLowerCase();

        if (tecnico && row.data('tecnico') !== tecnico) return false;
        if (oficina && row.data('oficina') !== oficina) return false;
        if (marca && row.data('marca') !== marca) return false;
        if (duplicado !== "" && row.data('duplicate').toString() !== duplicado) return false;
        if (searchText && !row.data('search').includes(searchText)) return false;

        return true;
    });

    $('#filterTecnico, #filterOficina, #filterMarca, #filterDuplicados').on('change', function () {
        table.draw();
    });

    $('#searchText').on('keyup', function () {
        table.draw();
    });
});

// Limpiar filtros
function limpiarFiltros() {
    $('#filterTecnico').val('');
    $('#filterOficina').val('');
    $('#filterMarca').val('');
    $('#filterDuplicados').val('');
    $('#searchText').val('');
    $('#inventoryTable').DataTable().draw();
}

// Confirmar eliminación
function confirmarEliminar(id, nombre) {
    $('#deleteEquipmentName').text(nombre);
    $('#deleteForm').attr('action', '/eliminar/' + id);
    var modal = new bootstrap.Modal(document.getElementById('deleteModal'));
    modal.show();
}

// Inicializar Select2
$(document).ready(function () {
    $('.select2').select2({ width: '100%' });
    $('#inventoryTable').DataTable({
        pageLength: 10,
        lengthChange: false,
        ordering: true
    });
});

// Confirmar eliminación
function confirmarEliminar(id, nombrePc) {
    $('#deleteEquipmentName').text(nombrePc);
    $('#deleteForm').attr('action', '/eliminar/' + id);
    var modal = new bootstrap.Modal(document.getElementById('deleteModal'));
    modal.show();
}

// Limpiar filtros
function limpiarFiltros() {
    $('.select2').val('').trigger('change');
    $('#searchText').val('');
    $('#inventoryTable').DataTable().search('').draw();
}

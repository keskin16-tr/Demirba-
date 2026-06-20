$(document).ready(function () {

    const table = $('#data_table').DataTable({
        order: [],
        autoWidth: false,
        lengthMenu: [[10, 25, 50, -1], [10, 25, 50, "Tümü"]],
        language: { url: "https://cdn.datatables.net/plug-ins/2.0.8/i18n/tr.json" },
        retrieve: true

        columnDefs: [
            { orderable: false, targets: 0 },
            { searchable: false, targets: 0 }
        ],

        dom:
            "<'row mb-2'<'col-md-6'B><'col-md-6'f>>" +
            "<'row'<'col-sm-12'tr>>" +
            "<'row'<'col-sm-5'i><'col-sm-7'p>>",

        buttons: [
            {
                extend: "colvis",
                text: "Sütunları Göster/Gizle",
                className: "btn btn-outline-secondary"
            }
        ]
    });

    table.buttons().container().appendTo("#dt-buttons-container");
    $("#data_table_filter").hide();

    $(".filter-input").on("keyup change", function () {
        const index = $(this).data("column-index");
        table.column(index).search(this.value).draw();
    });

    const selectedRows = new Set();

    $("#data_table").on("change", "input[name='selected_rows_checkbox']", function () {
        const val = $(this).val();
        this.checked ? selectedRows.add(val) : selectedRows.delete(val);
    });

    $("#select_all").on("click", function () {
        const checked = this.checked;

        table.rows({ search: "applied" }).nodes().to$()
            .find("input[name='selected_rows_checkbox']")
            .each(function () {
                this.checked = checked;
                const v = $(this).val();
                checked ? selectedRows.add(v) : selectedRows.delete(v);
            });
    });

    $("form").on("submit", function () {
        selectedRows.forEach(v => {
            $("<input>", {
                type: "hidden",
                name: "selected_rows",
                value: v
            }).appendTo("form");
        });
    });

});

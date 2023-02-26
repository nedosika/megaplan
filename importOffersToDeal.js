<input type="file" id="js-upload-file" />
<label for="js-upload-file" id="js-upload-file-label" class="btn-2">Импорт</label>

<style>
    #js-upload-file {
        height: 0;
        overflow: hidden;
        width: 0;
    }

    #js-upload-file + label {
        border: none;
        border-radius: 5px;
        color: #fff;
        cursor: pointer;
        display: inline-block;
        font-family: "Rubik", sans-serif;
        font-size: inherit;
        font-weight: 500;
        outline: none;
        padding: 5px 20px;
        position: relative;
        transition: all 0.3s;
        vertical-align: middle;
    }

    #js-upload-file + label.btn-2 {
        background-color: #497f42;
        border-radius: 5px;
        overflow: hidden;
    }
    #js-upload-file + label.btn-2:hover {
        background-color: #99c793;
    }

    #js-upload-file + label.error {
        background-color: #d9534f
    }

    #js-upload-file + label.error:hover {
        background-color: #FC9F8B
    }
</style>

<script>
    a9n.js("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js").then(
        () => {
            make_xlsx_lib(XLSX);
            const units = new Map();

            const uploadFileLabel = window.document.getElementById('js-upload-file-label');

            const inputFileLabelText = {
                defaultText: uploadFileLabel.innerText,
                loading: 'Загрузка файла...',
                uploading: 'Отправка на сервер...',
                error: 'Ошибка',
                done: 'Готово'
            }

            const sendFile = (data) => {
                uploadFileLabel.innerText = inputFileLabelText.uploading;

                const workbook = XLSX.read(data, {
                    type: 'binary'
                });

                workbook.SheetNames.forEach((sheetName) => {
                    const XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], {header:["A","B","C","D","E","F","G","H","I","J","K","L"], skipHeader: true});
                    const dealId = window.location.pathname.split('/')[2];
                    const positions = XL_row_object.map((position) => ({
                        contentType: "OfferRow",
                        deal: {
                             id: dealId,
                             contentType: "Deal"
                        },
                        quantity: position['D'] || 1,
                        price: {
                               contentType: "Money",
                               currency: "RUB",
                               value: position['J'] || 0,
                               valueInMain: position['J'] || 0,
                               rate: 1
                        }, 
                        //Category1000045CustomFieldBrakSht: 0,
                        name: position['C'],
                        //article: `${position['B'] || ''}`,
                        offer: {                        
                            contentType: "Offer",
                            id: 14,
                            //article: `${position['B'] || ''}`,
                            //unit: units.get(position['I']),
                            //description: `${position['L'] || ''}`,
                            //Category1000045CustomFieldBrakSht:  position['H'],
                            //Category1000045CustomFieldKolVoShtGotovo: position['E'],
                            //Category1000045CustomFieldKommentariy: String(position['L'] || ''),
                            //Category1000045CustomFieldNeobhodimaPokraska: !!position['F'],
                            //Category1000045CustomFieldPokrashenoSht: position['G'],
                            //price: {
                               //contentType: "Money",
                               //currency: "RUB",
                               //value: position['J'] || 0,
                               //valueInMain: position['J'] || 0,
                               //rate: 1
                            //},                
                       },
                    }));

                    console.log(positions)

                    const input = window.document.getElementById('js-upload-file');

                    $.ajax({
                        url: `/api/v3/deal/${dealId}`,
                        method: 'post',
                        data: JSON.stringify({
                            contentType: 'Deal',
                            positions,
                        }),
                        dataType: 'json',
                        contentType: 'application/json',
                        success: (d) => {
                            console.log(d)
                            uploadFileLabel.innerText = inputFileLabelText.done;
                            setTimeout(() => uploadFileLabel.innerText = inputFileLabelText.defaultText, 5000);
                            input.disabled = false;
                        },
                        error: function (xhr, ajaxOptions, thrownError) {
                            console.log(xhr, ajaxOptions, thrownError);
                            uploadFileLabel.innerText = inputFileLabelText.error;
                            uploadFileLabel.classList.add('error')
                            setTimeout(() => {
                                uploadFileLabel.innerText = inputFileLabelText.defaultText;
                                uploadFileLabel.classList.remove('error');
                                input.disabled = false;
                            }, 5000);
                        }
                    });
                })
            };

            const uploadFile = () => {
                const input = window.document.getElementById('js-upload-file');

                uploadFileLabel.innerText = inputFileLabelText.loading;
                input.disabled = true;

                if (typeof (FileReader) != "undefined") {
                    const reader = new FileReader();

                    //For Browsers other than IE.
                    if (reader.readAsBinaryString) {
                        reader.onload = function (e) {
                            sendFile(e.target.result);
                        };
                        reader.readAsBinaryString(input.files[0]);
                    } else {
                        //For IE Browser.
                        reader.onload = function (e) {
                            let data = "";
                            const bytes = new Uint8Array(e.target.result);
                            for (let i = 0; i < bytes.byteLength; i++) {
                                data += String.fromCharCode(bytes[i]);
                            }
                            sendFile(data)
                        };
                        reader.readAsArrayBuffer(input.files[0]);
                    }
                }
                else {
                    alert("This browser does not support HTML5.");
                }
            }

            const setUnits = ({data}) => data.forEach((unit) => units.set(unit.name, {contentType: "Unit", id: unit.id}));

            $('#js-upload-file').change(uploadFile);

            $.get('/api/v3/unit', setUnits);

            const filter = {
                // "fields": [
                //     "name",
                //     "type",
                //     "responsibles",
                //     "status",
                //     "countPositiveDeals",
                //     "countDeals",
                //     "summDeals",
                //     "summPositiveDeals"
                // ],
                "filter": {
                    "contentType": "CrmFilter",
                    "id": null,
                    "config": {
                        "contentType": "FilterConfig",
                        "termGroup": {
                            "contentType": "FilterTermGroup",
                            "join": "and",
                            "terms": [
                                {
                                    "contentType": "FilterTermRef",
                                    "field": "Category1000045CustomFieldUid",
                                    "comparison": "equals",
                                    "value": [1111]
                                }
                            ]
                        },
                        "filterId": null
                    }
                },
                "limit": 100
            };

            //$.get('/api/v3/offer?' + JSON.stringify(filter), (result) => console.log(result))
        });
</script>

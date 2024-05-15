sap.ui.define([
    "sap/ui/core/mvc/Controller",
    './xlsx/xlsx',
    './xlsx/xlsx.bundle'
],
    /**
     * @param {typeof sap.ui.core.mvc.Controller} Controller
     */
    function (Controller, XLSXjs, styleXLSXjs) {
        "use strict";

        return Controller.extend("zpa.controller.Main", {
            amountFields: [
                'tong',
                'amt001',
                'amt002',
                'amt003',
                'amt004',
                'amt005',
                'amt006',
                'amt007',
                'amt008',
                'amt009',
                'amt010',
                'amt011',
                'amt012'
            ],

            onInit: function () {

            },
            onInitial: function () {

            },
            onExport: function () {
                let thatController = this
                let oModel = this.getView().getModel()
                let filters = this.getView().byId('smartFilterBar').getFilters()
                let parameters = this.getView().byId('smartFilterBar').getFilterData()
                console.log(parameters)
                thatController._selectedMonth = []
                /* 
                * Get parameters
                * Lấy những tháng được chọn
                */
                filters.forEach((filter) => {
                    if (filter.sPath == 'fiscalperiod') {
                        if (filter.sOperator == 'EQ') {
                            thatController._selectedMonth.push(filter.oValue1)
                        } else {
                            for (let month = parseInt(filter.oValue1); month <= parseInt(filter.oValue2); month++) {
                                thatController._selectedMonth.push((`00${month}`).slice(-3))
                            }
                        }
                    }
                })
                this.selectedYear = parameters['$Parameter.FiscalYear']
                let paramsUrl = this.getView().byId('smartFilterBar').getParameterBindingPath()
                /**
                 * Get number of entries
                 */
                oModel.read(`${paramsUrl}/$count`, {
                    /* 
                    * Lấy số number of entries trước, sau đó thực hiện query data sau
                    */
                   
                    success: function (number) {
                        console.log(number)
                        oModel.read(`${paramsUrl}`, {
                            filters: filters,
                            urlParameters: {
                                '$top': number
                            },
                            success: function (data) {
                                thatController._data = new Map(data.results.map(i => [i.maso, i]));
                                thatController._data.forEach((node, key) => {
                                    if (node.formular && node.formular !== '' && !node.isCal) {
                                        thatController.passValueToFormula(node, key)
                                    }
                                })
                                thatController.exportExcel()

                            }
                        })
                    }
                })
            },
            passValueToFormula: function (node, key) {
                const regex = /(?<=<).+?(?=>)/g
                let varLst = node.formular.match(regex)
                this.amountFields.forEach((fieldname) => {
                    node[fieldname + "Final"] = `${node.formular}`
                })
                varLst.forEach((varKey) => {
                    let component = this._data.get(varKey)
                    if (varKey == key) {
                        this.amountFields.forEach((fieldname) => {
                            node[fieldname + "Final"] = node[fieldname + "Final"].replaceAll(`<${varKey}>`, node[fieldname])
                        })
                    } else {
                        /**
                         * Có formula và đã tính kết quả
                         * Không có formula
                         */
                        if ((component.formular && component.formular !== '' && component.isCal) ||
                            (component.formular == '') ||
                            (!component.formular)) {
                        } else {
                            /**
                             * Có formula và chưa tính kết quả
                             */
                            this.passValueToFormula(component, component.maso)
                        }
                        this.amountFields.forEach((fieldname) => {
                            console.log(varKey, component[fieldname])
                            node[fieldname + "Final"] = node[fieldname + "Final"].replaceAll(`<${varKey}>`, component[fieldname])
                        })
                    }
                })
                this.amountFields.forEach((fieldname) => {
                    node[fieldname] = eval(node[fieldname + "Final"])
                })
                node.isCal = true
            },
            convertExcelColCharacter: function(index){
                var result = '';
                do {
                    result = (index % 26 + 10).toString(36) + result;
                    index = Math.floor(index / 26) - 1;
                } while (index >= 0)
                return result.toUpperCase();
            },          
            cellStyle: function( textStyleFont,  alignmentHorizontal ){
                return {
                    font: {
                        bold: textStyleFont.isBold,
                        name: "Times New Roman"
                    },
                    alignment: {
                        horizontal: alignmentHorizontal
                    },
                    border: {
                        top: { style: "thin", color: {rgb:"000000"}},
                        bottom: { style: "thin", color: {rgb:"000000"}},
                        left: { style: "thin", color: {rgb:"000000"}},
                        right: { style: "thin", color: {rgb:"000000"}}
                    }
                }
            },  
            exportExcel: function () {
                /* Prepare column list */
                const VND = new Intl.NumberFormat('en-DE');
                let listColMapping = []
                let excelData = []

                listColMapping.push({ name: "Đơn vị", colField: "donvi"})
                listColMapping.push( { name: "Mã data", colField : "madata"} )
                listColMapping.push( { name: "Mã số", colField : "maso"} )
                listColMapping.push( { name: "Chỉ tiêu", colField : "nodedes"} )
                listColMapping.push( { name: `Năm ${this.selectedYear}`, colField: "tong", type:'currency'})
                this._selectedMonth.forEach((month)=>{
                    listColMapping.push( { name: "Tháng " + month, colField: `amt${month}`, type:'currency'})
                })
                listColMapping.push( { name: "Xuất Nhập Khẩu", colField : "xnk", type:'currency'} )
                listColMapping.push( { name: "Dịch vụ Thương Mại", colField : "dvtm", type:'currency'} )
                listColMapping.push( { name: "Kinh doanh Thương Mại", colField : "kdtm", type:'currency'} )
                listColMapping.push( { name: "Thầu", colField : "thau", type:'currency'} )
                listColMapping.push( { name: "Kho vận", colField : "khovan", type:'currency'} )
                listColMapping.push( { name: "Tài chính - Kế toán", colField : "tckt", type:'currency'} )
                listColMapping.push( { name: "Quản lý chất lượng", colField : "qlcl", type:'currency'} )
                listColMapping.push( { name: "Tổ chức nhân sự", colField : "tcns", type:'currency'} )
                listColMapping.push( { name: "Công nghệ thông tin", colField : "cntt", type:'currency'} )
                listColMapping.push( { name: "Hành chính", colField : "hc", type:'currency'} )
                listColMapping.push( { name: "Quản lý dự án", colField : "qlda", type:'currency'} )
                listColMapping.push( { name: "Ban điều hành", colField : "bdh", type:'currency'} )
                listColMapping.push( { name: "Văn phòng", colField : "vp", type:'currency'} )
                console.log(listColMapping)

                var headerStyle = {
                    font: {
                        bold: true,
                        name: "Times New Roman"
                    },
                    alignment: {
                        horizontal: "center"
                    },
                    border: {
                        top: { style: "thin", color: {rgb:"000000"}},
                        bottom: { style: "thin", color: {rgb:"000000"}},
                        left: { style: "thin", color: {rgb:"000000"}},
                        right: { style: "thin", color: {rgb:"000000"}}
                    }
                }

                /* Append header title */
                let row = []
                let rowIndex = 1
                let excelStyle = []
                excelData.push(['BÁO CÁO THỰC HIỆN'])              
                rowIndex += 1

                excelData.push([])
                rowIndex += 1

                
                listColMapping.forEach((col, index)=>{
                    row.push(col.name)
                    excelStyle.push({
                        cell : `${this.convertExcelColCharacter(index)}${rowIndex}`,
                        style: headerStyle
                    })
                })
                excelData.push(row)
                rowIndex += 1
                /* Append data */
                const colNum = listColMapping.length
                /* 
                * Tạo array of array data để export excel
                * [ 
                *   [A1, B1, C1, .. AA1] , 
                *   [A2, B2, C2, .. AA2], 
                *   [A3, B3, C3, .. AA3] 
                * ]  
                * excelStyle: array chứa style của các cells
                * excelData : array chứa data của các cells
                */
                this._data.forEach((value, key)=>{
                    if (value.hidden) {
                        return
                    }
                    row = []
                    listColMapping.forEach((col, index)=>{
                        if (col.type == 'currency') {
                            value[col.colField] = VND.format(value[col.colField] ? value[col.colField] : 0)
                            excelStyle.push({
                                cell : `${this.convertExcelColCharacter(index)}${rowIndex}`,
                                style: {
                                    font: {
                                        name: "Times New Roman"
                                    },
                                    alignment: {
                                        horizontal: "center"
                                    },
                                    border: {
                                        top: { style: "thin", color: {rgb:"000000"}},
                                        bottom: { style: "thin", color: {rgb:"000000"}},
                                        left: { style: "thin", color: {rgb:"000000"}},
                                        right: { style: "thin", color: {rgb:"000000"}}
                                    }
                                }
                            })
                        } else {
                            excelStyle.push({
                                cell : `${this.convertExcelColCharacter(index)}${rowIndex}`,
                                style: {
                                    font: {
                                        name: "Times New Roman"
                                    },
                                    border: {
                                        top: { style: "thin", color: {rgb:"000000"}},
                                        bottom: { style: "thin", color: {rgb:"000000"}},
                                        left: { style: "thin", color: {rgb:"000000"}},
                                        right: { style: "thin", color: {rgb:"000000"}}
                                    }
                                }
                            })
                        }
                        row.push(value[col.colField] ? value[col.colField] : '')

                    })
                    excelData.push(row)
                    rowIndex += 1
                })

                var xlsxData = XLSX.utils.aoa_to_sheet(excelData)
                const spreadsheet = XLSX.utils.book_new()
                XLSX.utils.book_append_sheet(spreadsheet, xlsxData, 'Data')
                excelStyle.forEach(value=>{
                    spreadsheet.Sheets["Data"][value.cell].s = value.style
                })
                spreadsheet.Sheets['Data'].A1.s =  {
                    font: {
                        name: "Times New Roman",
                        bold: true,
                        color: {rgb:'E54121'}
                    }
                 }
                XLSX.writeFile(spreadsheet, "Báo cáo thực hiện.xlsx");
            }
        });
    });
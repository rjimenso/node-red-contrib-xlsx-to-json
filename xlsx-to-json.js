module.exports = function(RED){
    function convertxlsx(config){
        RED.nodes.createNode(this,config);
        //Los inputs
        this.filepath = config.filepath;
        this.rangecell = config.rangecell;
        this.columnkey = config.columnkey;
        this.sheet = config.sheet;

        var node = this;
        node.on('input', function(msg){
            const excelToJson = require('convert-excel-to-json');
            let result = {};
            let columnas;
            const defaultColumn = '{"*": "{{columnHeader}}"}';

            if (!node.rangecell){
                node.rangecell = msg.rangecell;
            }

            if(!node.columnkey){
                node.columnkey = msg.columnkey;
            }

            if (!node.sheet){
                node.sheet = msg.sheet;
            }

            if(node.columnkey){
                columnas = JSON.parse(node.columnkey);
            } else {
                columnas = JSON.parse(defaultColumn);
            } 

            if (msg.hasOwnProperty('payload') && Buffer.isBuffer(msg.payload)) {
                result = excelToJson({
                    source: Buffer.from(msg.payload),
                    range: node.rangecell,
                    columnToKey: columnas
                // sheets: node.sheet
                })
            } else {      
                if(!node.filepath){
                    node.filepath = msg.filepath;
                }
                result = excelToJson({
                    sourceFile: node.filepath,
                    range: node.rangecell,
                    columnToKey: columnas
                // sheets: node.sheet
                })
            }
            if (node.sheet){
                msg.payload = result[node.sheet];
            }else{
                msg.payload = result;
            }
            node.send(msg);
        });
    }
    RED.nodes.registerType("XLSX-to-json", convertxlsx);
}
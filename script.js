var xlsx = require('node-xlsx');
var fs = require('fs');
var filesFolder = __dirname + '/files';
var resultFolder = __dirname + '/results';
let files = fs.readdirSync( filesFolder ).filter( function( file ) {return file.match( /.*\.(xls|xlsx)/ig);});

files.forEach(function (item, order) {
    file = xlsx.parse( filesFolder+'/'+item );
    var writeStr = "";
    var filename = resultFolder+'/honda_'+order+'.csv';

    file.forEach(function (page) {
        if( page.name != 'Proposta' )
            return false;

        page.data.forEach(function (context,line) { 
            if( line <= 3 )
            {
                return false;                  
            }
                
            writeStr += handlingArray( context );
            if( line % 200 == 0 )
            {   
                var filename = resultFolder+'/'+item+'-file_'+line+'.csv';
                fs.writeFile( filename, writeStr, function(err) {
                    if(err) {
                        return console.log(err);
                    }
                    console.log(filename+" foi gerado até a linha "+line+" !");
                });              
                writeStr = "";
            }
            
        })     
    })
})

function handlingArray(array)
{
    array.splice(0,1);
    return array.join(";")+"\n";
}


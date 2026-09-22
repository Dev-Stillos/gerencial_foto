$(document).ready(function(){


    $('#content').html('<div id="timeline" style="width: 100%; height: 550px"></div>');
    google.setOnLoadCallback(drawChart);
    
    $.post('includes/asp/ple_lista_projetos.asp?date='+Date(),function(data){
        		 $('#content').append(data);
	 
	 });

    
});



function drawChart() {
  var container = document.getElementById('timeline');

  var chart = new google.visualization.Timeline(container);

  var dataTable = new google.visualization.DataTable();
  dataTable.addColumn({ type: 'string', id: 'Projeto' });
  dataTable.addColumn({ type: 'string', id: 'Descrição' });
  dataTable.addColumn({ type: 'date', id: 'Inicio' });
  dataTable.addColumn({ type: 'date', id: 'fim' });

  $.post('includes/asp/ple_timeline.asp?date='+Date(),function(data){
    console.log(data)
    data = data.split(']][[');
    for (i=0;i<data.length;i++){
        projeto = data[i];
        projeto = projeto.replace('[[','').replace(']]','');
        projeto = projeto.split('::');
        console.log(projeto);
        data1 = projeto[2].split("/");
        data2 = projeto[3].split("/"); 
        dataTable.addRows([
           [  projeto[0], projeto[1], new Date(data1[2],data1[1]-1,data1[0]), new Date(data2[2],data2[1]-1,data2[0]) ],  
        ]);
        
        var options = {
            timeline: { colorByRowLabel: true }
        };
        
    	chart.draw(dataTable, options);
    }    
     
  });

  /*
 dataTable.addRows([
    [  'Solicitações', new Date(2014, 2, 17), new Date(2014, 6, 31) ],
    [  'Controle de Estoque', new Date(2014, 2, 17), new Date(2014, 3, 31) ],
    [  'Gestão de contratos',       new Date(2014, 4, 1),   new Date(2014, 4, 10) ],
    [  'Sistema Gerêncial',      new Date(2014, 4, 10),   new Date(2014, 5, 30) ],
    [  'Integração - Sistemas',      new Date(2014, 6, 1),   new Date(2014, 7, 23) ],    
    ]);
*/

}



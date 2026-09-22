<!DOCTYPE html>
<html lang="pt-br">
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<title>Pré Pedido</title>
	<link rel="stylesheet" type="text/css" href="css/normalize.css"></link>
	<link rel="stylesheet" type="text/css" href="css/bootstrap/css/bootstrap.min.css"></link>
	<link rel="stylesheet" type="text/css" href="css/bootstrap/css/bootstrap-theme.min.css"></link>
	<link rel="stylesheet" type="text/css" href="css/style-pre-pedido.css"></link>
</head>
<body>
	<div class="container-fluid">
		<div class="container">
			<div class="row">
				<div class="col-md-12">
					<form action="#" method="post" accept-charset="utf-8">
						<div class="form-group col-md-2">
							<label for="data_evento">Data do evento</label>
							<input type="text" id="data_evento" name="data_evento" value="" class="form-control">
						</div>
						
						<div class="form-group col-md-2">
							<label for="evento_tipo">Evento tipo</label>
							<select name="evento_tipo"  id="evento_tipo" class="form-control">
								<option value="colacao">Colação</option>
								<option value="baile">Baile</option>
							</select>
						</div>

						<div class="form-group col-md-8">
							<label for="local_evento">Local do evento</label>
							<input type="text" id="local_evento" name="local_evento" value="" class="form-control">
						</div>							
						<div class="form-group col-md-12">
							<div class="col-md-2">
								<label for="evento_tipo">Contrato</label>
								<div class="input-group">
								  <input type="text" class="form-control" aria-label="Contrato" class="form-control">
								  <span class="input-group-addon"><button type="button">+</button></span>
								</div>	
							</div>
							<div class="col-md-10">
								<table class="table table-bordered table-hover">
									<thead>
										<tr>
											<th>Contrato</th>
											<th>Instituição</th>
											<th>Curso</th>
										</tr>
									</thead>
									<tbody>
										<tr>
											<td><input type="text" name="contrato" value=""></td>
											<td><input type="text" name="instituicao" value=""></td>
											<td><input type="text" name="curso" value=""></td>
										</tr>
									</tbody>
								</table>	
							</div>
						</div>
						<div class="form-group col-md-2">
							<label for="horario">Horário do evento</label>
							<input type="text" id="horario" name="horario" value="" class="form-control">
						</div>	
						<div class="form-group col-md-2">
							<label for="formandos">Número de formandos</label>
							<input type="number" id="formandos" name="formandos" value="" class="form-control" min="0" max="10000">
						</div>	
						<div class="form-group col-md-8">
							<label for="coordenador">Nome do Coordenador</label>
							<input type="text" id="coordenador" name="coordenador" value="" class="form-control">
						</div>
						<div class="col-md-12">
							<button class="btn btn-success btn-block" type="submit">Enviar</button>
						</div>	
					</form>
				</div>
			</div>
		</div>
	</div>
</body>
</html>
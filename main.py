#!/usr/bin/env python3
import click
import sys
from pathlib import Path
from src.transformer import ExcelTransformer


@click.command()
@click.option('--input', '-i', required=True, help='Archivo de entrada (HTML, XLS, XLSX, CSV)')
@click.option('--config', '-c', required=True, help='Archivo de configuración YAML')
@click.option('--output', '-o', help='Archivo de salida (opcional, usa config si no se especifica)')
@click.option('--verbose', '-v', is_flag=True, help='Mostrar información detallada')
def main(input, config, output, verbose):
    """
    Transformador de archivos Excel basado en configuración.
    
    Ejemplos:
    
    python main.py -i reporte.xls -c configs/balance_sheet.yaml
    
    python main.py -i datos.xlsx -c configs/import_reporting.yaml -o resultado.txt
    """
    
    try:
        # Verificar que los archivos existen
        if not Path(input).exists():
            click.echo(f"Error: Archivo de entrada no encontrado: {input}", err=True)
            sys.exit(1)
            
        if not Path(config).exists():
            click.echo(f"Error: Archivo de configuración no encontrado: {config}", err=True)
            sys.exit(1)
        
        if verbose:
            click.echo(f"Cargando archivo: {input}")
            click.echo(f"Aplicando configuración: {config}")
        
        # Crear transformador y procesar
        transformer = ExcelTransformer(config)
        result_path = transformer.process(input, output)
        
        if verbose:
            click.echo(f"Datos cargados: {len(transformer.df)} filas, {len(transformer.df.columns)} columnas")
            click.echo(f"Transformaciones aplicadas: {len(transformer.config.get('transformations', []))}")
        
        click.echo(f"Archivo procesado exitosamente: {result_path}")
        
    except Exception as e:
        click.echo(f"Error procesando archivo: {str(e)}", err=True)
        if verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


@click.group()
def cli():
    """Herramientas para transformar archivos Excel."""
    pass


@cli.command()
def list_configs():
    """Lista las configuraciones disponibles."""
    configs_dir = Path("configs")
    if not configs_dir.exists():
        click.echo("Directorio 'configs' no encontrado")
        return
    
    click.echo("Configuraciones disponibles:")
    for config_file in configs_dir.glob("*.yaml"):
        click.echo(f"  - {config_file.stem}")


@cli.command()
@click.argument('config_name')
def show_config(config_name):
    """Muestra el contenido de una configuración."""
    config_path = Path(f"configs/{config_name}.yaml")
    if not config_path.exists():
        click.echo(f"Configuración no encontrada: {config_name}")
        return
    
    click.echo(f"Configuración: {config_name}")
    click.echo("=" * 50)
    with open(config_path, 'r', encoding='utf-8') as f:
        click.echo(f.read())


if __name__ == '__main__':
    # Si se ejecuta directamente, usar el comando principal
    if len(sys.argv) == 1:
        cli()
    else:
        # Si hay argumentos, intentar el comando principal primero
        try:
            main(standalone_mode=False)
        except click.exceptions.MissingParameter:
            # Si falta parámetro, mostrar ayuda del grupo
            cli()
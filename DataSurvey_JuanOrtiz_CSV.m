%DESARROLLADOR: JUAN ORTIZ (2025)
%EMPRESA: CONSORCIO SIG-ELECTRIC
%DESCRIPCION: Script para generar archivos CSV para ser utilizados en un
%formulario de ArcGIS Survey123
%VERSION 2

clc
clear 
format LONGG

%%                   CAMBIOS IMPORTANTES 
%Estos cambios se deben hacer en caso que se reemplace el archivo de
%ingreso o salida de datos, cambios de variable y hojas en excel

%%%%%%%%%%%%%%%%%%%%%% INGRESAR DATOS %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ExcelInputBase = 'TABLA GENERAL 59E.xlsx';
folderOut = '59ECSV'; % Carpeta donde se guardarán los archivos CSV

%%%%%%%%%%%%%%%%%%%%% SALIDA DE DATOS %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if ~exist(folderOut, 'dir')
    mkdir(folderOut); % Crear la carpeta si no existe
end
%%%%%%%%%%%%%%%%%%% PREPARANDO EL AMBIENTE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Hojas de lectura
sheetNames = struct(...
    'TRAFO', 'TRAFO', ...
    'POSTE', 'POSTE', ...
    'PUNTO_CARGA', 'PUNTO DE CARGA', ...
    'MEDIDORES', 'MEDIDORES' ...
);

% Columnas y cabeceras
headers = struct(...
    'TRAFO', {{"name", "label", "alimentador"}}, ...
    'POSTE', {{"name", "label", "trafos", "alimentador"}}, ...
    'MEDIDORES', {{"name", "label", "acometidas", "alimentador"}} ...
);

% Patrones para extraer información
patternDigits = '(?<=1400)\d{2}(?=.*[A-Z](?!.*[A-Z]))'; % Dos dígitos específicos
patternLastLetter = '[A-Z](?!.*[A-Z])'; % Última letra mayúscula

%%PROCESAMIENTO GENERAL
% Lista de tablas y configuraciones
tablesConfig = {
    'Trafos', sheetNames.TRAFO, headers.TRAFO;
    'Postes', sheetNames.POSTE, headers.POSTE;
    'MedidoresSA', sheetNames.MEDIDORES, headers.MEDIDORES;
};


for i = 1:size(tablesConfig, 1)
    tableName = tablesConfig{i, 1};
    sheetName = tablesConfig{i, 2};
    header = tablesConfig{i, 3};

    fprintf('********************************Procesando %s...\n', tableName);

    % Leer datos de la hoja correspondiente
    dataTable = leerDatosExcel(ExcelInputBase, sheetName);

        % Eliminar filas con valores vacíos o '0' en la columna principal
    if ismember('CODIGOPUES', dataTable.Properties.VariableNames)
        dataTable(strcmp(dataTable.CODIGOPUES, '0'), :) = [];
        dataTable(cellfun(@isempty, dataTable.CODIGOPUES), :) = [];
    elseif ismember('OID_POSTE', dataTable.Properties.VariableNames)
        dataTable(strcmp(dataTable.OID_POSTE, '0'), :) = [];
        dataTable(cellfun(@isempty, dataTable.OID_POSTE), :) = [];
    elseif ismember('MDENUMEMP', dataTable.Properties.VariableNames)
        dataTable(strcmp(dataTable.MDENUMEMP, '0'), :) = [];
        dataTable(cellfun(@isempty, dataTable.MDENUMEMP), :) = [];
    end

        % Procesar columna ALIMENTADO si existe
    if ismember('ALIMENTADO', dataTable.Properties.VariableNames)

        es1400 = startsWith(dataTable.ALIMENTADO, '1400');
        dataTable.ALIMENTADOR = dataTable.ALIMENTADO; % Inicializar con valores originales

        if any(es1400)
           fprintf('Se encontraron filas que comienzan con "1400".\n');

           % Extraer patrones
            digits = regexp(dataTable.ALIMENTADO(es1400), patternDigits, 'match');
            lastLetter = regexp(dataTable.ALIMENTADO(es1400), patternLastLetter, 'match');

            % Combinar resultados
            dataTable.ALIMENTADOR(es1400) = cellfun(@(d, l) [d{:} l{:}], digits, lastLetter, 'UniformOutput', false);
        else
             fprintf('No se encontraron filas que comiencen con "1400".\n');
        end 
    else
        warning('La columna ALIMENTADO no existe en la tabla %s.\n', tableName);
    end  

 % Crear la tabla final
    if strcmp(tableName, 'Trafos')
        % Detectar y eliminar duplicados en 'CODIGOPUES' y 'ALIMENTADOR'
        [~, uniqueIdx] = unique(dataTable(:, {'CODIGOPUES', 'ALIMENTADOR'}), 'rows', 'stable');
        dataTable = dataTable(uniqueIdx, :); % Conservar solo los registros únicos
        
         % Crear la tabla final
        finalTable = [header; dataTable.CODIGOPUES, dataTable.CODIGOPUES, dataTable.ALIMENTADOR];

    elseif strcmp(tableName, 'Postes')
        % Detectar y eliminar duplicados en 'OID_POSTE' y 'ALIMENTADOR'
        [~, uniqueIdx] = unique(dataTable(:, {'OID_POSTE', 'ALIMENTADOR'}), 'rows', 'stable');
        dataTable = dataTable(uniqueIdx, :); % Conservar solo los registros únicos
        
        % Crear la tabla final
        finalTable = [header; dataTable.OID_POSTE, dataTable.OID_POSTE, dataTable.TRAFO, dataTable.ALIMENTADOR];

    elseif strcmp(tableName, 'MedidoresSA')
        % Detectar y eliminar duplicados en 'MDENUMEMP' y 'ALIMENTADOR'
        [~, uniqueIdx] = unique(dataTable(:, {'MDENUMEMP', 'ALIMENTADOR'}), 'rows', 'stable');
        dataTable = dataTable(uniqueIdx, :); % Conservar solo los registros únicos
        
        % Crear la tabla final
        finalTable = [header; dataTable.MDENUMEMP, dataTable.MDENUMEMP, dataTable.PC, dataTable.ALIMENTADOR];
    end

    % Guardar el archivo CSV
    outputFileName = fullfile(folderOut, sprintf('%sCSV.csv', tableName));
    writecell(finalTable, outputFileName);
    fprintf('Archivo %s generado exitosamente.\n', outputFileName);
end 

fprintf('****************Archivos CSV GENERADOS **********************\n');
%%                     LECTURA DE DATOS  


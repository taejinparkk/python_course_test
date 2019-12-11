% time series generation for ratings data from BB

clc;
clear;

% file name of xls files containing ratings start with 
file_in   = 'ratings_banks_Adrian';
% result files start with the following prefix to disdinguis them form
% other users files
file_out  = 'banks_Adrian';

% type reduce_level = 'd'; for daily, otherwise for 'b' you get businessdays
reduce_level = 'q';  % w,m,q,y are also possible values

% which sheet in Names.xls contains the list of entities to look for
names = 'bank list Adrian';

% alter the following lines if you want to look for particular rating types
% otherwise leave types = {};
types = { 'Fitch Individual Rating', 'Fitch Outlook', 'Fitch LT Issuer Default Rating', 'Fitch LT FC Issuer Default', ...
            'Fitch Viability', ...
            'Moody''s Bank Financial Strength', 'Moody''s FC Curr Issuer Rating',  'Moody''s Foreign LT Bank Deposits', ...
            'Moody''s Issuer Rating', 'Moody''s Long Term Rating', 'Moody''s Long Term Bank Deposits', ...
            'Moody''s Outlook', 'Moody''s Senior Unsecured Debt', ...
            'S&P Outlook', 'S&P LT Foreign Issuer Credit' };

% types = {'Fitch Foreign Currency LT Debt', 'Fitch Local Currency LT Debt', ...
%          'Moody''s Foreign Currency LT Debt', 'Moody''s Local Currency LT Debt', ...
%          'S&P Foreign Currency LT Debt', 'S&P Local Currency LT Debt'};
% NOTE: ADDING A NEW TYPE requires amending the spreadsheets Fitch Ratings
% Types.xls, Moody's Ratings Types.xls or S&P Ratings Types.xls in the 
% sub directory conversion tables for rating types

% root path
path      = '\\Msfsshared\med\DRS\Databases\Ratings\';  
% data path 
path_data = [ path 'data\'];
% results path
path_res  = [ path 'results\Adrian\'];

% time series start
start_date = '1990-01-01';   % if '' the default '1990.01.01' is taken - the date must be the beginning of the download 

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

addpath('\\Msfsshared\med\DRS\Databases\Ratings\functions');

% I am getting the ratings data from xls files named ratings*.xls and store them in one big cell 
[ data, all_entities ] = get_ratings_data_NEW(path_data, file_in);   % note: entities are now security, as they do not change as often 
                                                                     % in BB as names
%%%%%%% what rating agencies %%%%%%%

rating_agencies = unique(data(:,4));

for k=1:size(rating_agencies,1)
   
    Index = strcmp(data(:,4), rating_agencies{k,:});
    
    data1 = data(Index,[1,2,3,5,6]);
    
    %%%%%%% what rating types %%%%%%%
    
    rating_types = unique(data1(:,3));
    
    for l=1:size(rating_types,1)
        
        Index  = strcmp(types, [rating_agencies{k} ' ' rating_types{l}] );
        
        % if types is empty I will do all rating types, if types is not
        % empty I will look only for the ones which are in types
        if( isempty( types ) | sum(Index) ) 
            
            [rating_agencies{k} ' ' rating_types{l}]
            
            Index  = strcmp( data1(:,3), rating_types{l,:});
            
            data2 = data1( Index, [1,2,4,5] );
            
            %%%%%%% what companies %%%%%%%
            
            companies = upper(unique( data2(:,1))); 
            
            ts_ratings = [];
            ts_watch   = [];
            
            for m=1:size(companies,1)
                
                Index  = strcmp( upper(data2(:,1)), companies{m,:});
                            
                data3 = data2( Index, [2,3,4] );
                
                % now I am creating a time series for companies{m,:} for rating_agencies{k} and rating_types{l}
                % I am extracting ratings and watches 
                
                [ts1, ts2] = make_ts(data3, path, rating_agencies{k}, rating_types{l}, companies{m,:}, start_date);
                
                if( strcmp( reduce_level, 'b' ) )
                    ts1 = reduceLevel(ts1,reduce_level,true);
                    ts2 = reduceLevel(ts2,reduce_level,true);
                elseif( strcmp( reduce_level, 'w' ) | strcmp( reduce_level, 'm' ) | strcmp( reduce_level, 'q' ) | strcmp( reduce_level, 'y' ) )
                    ts1 = ts1.compress(reduce_level,Aggr.last,true);
                    ts2 = ts2.compress(reduce_level,Aggr.last,true);
                elseif( ~strcmp( reduce_level, 'd' ) )
                    error( 'reduce level wrongly specified: possible values d, b, w, m, q, y');
                end
                
                ts_ratings = [ts_ratings ts1];
                ts_watch   = [ts_watch ts2];
                
            end
            
            [sorted_ts_ratings, short_names] = sort_by_Banks_NEW(ts_ratings, names, all_entities, reduce_level);
            [sorted_ts_watch, short_names]   = sort_by_Banks_NEW(ts_watch, names, all_entities, reduce_level); 
            
            writeExcel(sorted_ts_ratings, [path_res file_out '_' reduce_level '_' rating_agencies{k} ' ' rating_types{l} '.xlsx'], 'Sheet1','StartingCell' , 'A1');
            pause(2);
            writeExcel(sorted_ts_watch, [path_res file_out '_' reduce_level '_' rating_agencies{k} ' ' rating_types{l} ' WATCH.xlsx'], 'Sheet1', 'StartingCell' , 'A1');
            pause(2);
            % overwriting the heading with the bank name
            xlswrite([path_res file_out '_' reduce_level '_' rating_agencies{k} ' ' rating_types{l} '.xlsx'], short_names','Sheet1','B1');
            pause(2);
            xlswrite([path_res file_out '_' reduce_level '_' rating_agencies{k} ' ' rating_types{l} ' WATCH.xlsx'], short_names','Sheet1','B1');

        end
        
    end
    
end


%str=
%newStr = strrep(str,old,new);

z = load('B0030.mat');

len = length(z.B0030.cycle);

disp(len);
c=0;
d=0;
e=0;
for i = 1:len
    if strcmp( z.B0030.cycle(i).type, 'charge' )
        c=c+1;
        
        col_header={'Voltage_measured','Current_measured','Temperature_measured','Current_charge', 'Voltage_charge', 'Time' };     %Row cell array (for column labels)
        %row_header(1:10,1)={'Time'};     %Column cell array (for row labels)
        xlswrite('Charge.xlsx',col_header,c,'A1');     %Write data
    
        
        %[reqStruct1] = z.B0030.cycle(c).data;
        [reqStruct1] = z.B0030.cycle(i).data.Voltage_measured;
        A= transpose(reqStruct1);
        xlswrite('Charge.xlsx',A,c,'A2')
        
        
        [reqStruct2] = z.B0030.cycle(i).data.Current_measured;
        B= transpose(reqStruct2);
        xlswrite('Charge.xlsx',B,c,'B2');
      

        [reqStruct3] = z.B0030.cycle(i).data.Temperature_measured;
        C= transpose(reqStruct3);
        xlswrite('Charge.xlsx',C,c,'C2');

        [reqStruct4] = z.B0030.cycle(i).data.Current_charge;
        D= transpose(reqStruct4);
        xlswrite('Charge.xlsx',D,c,'D2');


       [reqStruct5] = z.B0030.cycle(i).data.Voltage_charge;
       E= transpose(reqStruct5);
       xlswrite('Charge.xlsx',E,c,'E2');

       [reqStruct6] = z.B0030.cycle(i).data.Time;
       F= transpose(reqStruct6);
       xlswrite('Charge.xlsx',F,c,'F2'); 
    
    elseif strcmp( z.B0030.cycle(i).type, 'discharge' )
        d=d+1;
        
        col_header={'Voltage_measured','Current_measured','Temperature_measured','Current_load', 'Voltage_load', 'Time', 'Capacity' };     %Row cell array (for column labels)
        %row_header(1:10,1)={'Time'};     %Column cell array (for row labels)
        xlswrite('Discharge.xlsx',col_header,d,'A1');     %Write data
    
        
        %[reqStruct1] = z.B0030.cycle(c).data;
        [reqStruct1] = z.B0030.cycle(i).data.Voltage_measured;
        A= transpose(reqStruct1);
        xlswrite('Discharge.xlsx',A,d,'A2')
        
        
        [reqStruct2] = z.B0030.cycle(i).data.Current_measured;
        B= transpose(reqStruct2);
        xlswrite('Discharge.xlsx',B,d,'B2');
      

        [reqStruct3] = z.B0030.cycle(i).data.Temperature_measured;
        C= transpose(reqStruct3);
        xlswrite('Discharge.xlsx',C,d,'C2');

        [reqStruct4] = z.B0030.cycle(i).data.Current_load;
        D= transpose(reqStruct4);
        xlswrite('Discharge.xlsx',D,d,'D2');


       [reqStruct5] = z.B0030.cycle(i).data.Voltage_load;
       E= transpose(reqStruct5);
       xlswrite('Discharge.xlsx',E,d,'E2');

       [reqStruct6] = z.B0030.cycle(i).data.Time;
       F= transpose(reqStruct6);
       xlswrite('Discharge.xlsx',F,d,'F2'); 
       
       xlswrite('Capacity.xls','Capacity','A1');
       [reqStruct7] = z.B0029.cycle(i).data.Capacity;
        G= reqStruct7;
        xlsappend('Capacity.xls',G,2); 
    else
        
        e=e+1;
        
        col_header={'Sense_current','Battery_current','Current_ratio','Battery_impedence', 'Rectified_impedence', 'Re', 'Rct' };     %Row cell array (for column labels)
        %row_header(1:10,1)={'Time'};     %Column cell array (for row labels)
        xlswrite('Impedance.xlsx',col_header,e,'A1');     %Write data
              
        [reqStruct1] = z.B0030.cycle(i).data.Sense_current;
        p = num2cell(reqStruct1);
        k=transpose(p);
        towrite = cellfun(@num2str , k, 'UniformOutput', false);
        xlswrite('Impedance.xlsx',towrite, e, 'A2')
       % xlswrite('Impedance.xlsx',X,e,'A2')
        
        
               
        
        %%%%%%%%%%% change in below lines as above%%%%%%%%%%%%%%%%%%               
        [reqStruct2] = z.B0030.cycle(i).data.Battery_current;
        
        B= transpose(reqStruct2);
        xlswrite('Impedance.xlsx',B,e,'B2');
      

        

        [reqStruct3] = z.B0030.cycle(i).data.Current_ratio;
        C= transpose(reqStruct3);
        xlswrite('Impedance.xlsx',C,e,'C2');

        [reqStruct4] = z.B0030.cycle(i).data.Battery_impedance;
        p = num2cell(reqStruct4);
        D= transpose(p);
        xlswrite('Impedance.xlsx',D,e,'D2');

      
       [reqStruct5] = z.B0030.cycle(i).data.Rectified_Impedance;
       p = num2cell(reqStruct5);
       E= transpose(p);
       xlswrite('Impedance.xlsx',E,e,'E2');

       [reqStruct6] = z.B0030.cycle(i).data.Re;
       F= transpose(reqStruct6);
       xlswrite('Impedance.xlsx',F,e,'F2'); 
       
       [reqStruct7] = z.B0030.cycle(i).data.Rct;
       G= transpose(reqStruct7);
       xlswrite('Impedance.xlsx',G,e,'G2'); 
        
        
        
        
        
        
        
        
        
        
        
        
       
    end
    
    
    
    
end
% disp(c);
% 
% len1=count( z.B0030.cycle(1).type, 'charge' ) ;
% disp(len1);
% a = zeros( len, 1 );
% %a=zeros(3,1);
% j=0;
% k=0;
% f=0;
% if any(strcmp(C, 'charge'))
%   disp('charge found');
% end
% for i = 1:len
%     if strcmp( z.B0030.cycle(i).type, 'charge' ) 
%         
%        % a(ii) = z.B0030.cycle(ii).data.Capacity;
%         [reqStruct1] = z.B0030.cycle(i).data.Voltage_measured;
%         A= transpose(reqStruct1);
%         xlswrite('Charge.xlsx',A,i+j,'A1');
%         
%         [reqStruct2] = z.B0030.cycle(i).data.Current_measured;
%         B= transpose(reqStruct2);
%         xlswrite('Charge.xlsx',B,i+j,'B1');
% 
%         [reqStruct3] = z.B0030.cycle(i).data.Temperature_measured;
%         C= transpose(reqStruct3);
%         xlswrite('Charge.xlsx',C,i+j,'C1');
%         
%         
%         %a(ii) = z.B0030.cycle(i).data.Voltage_measured;
%         %disp(a(i));
%         j=j+1;
%     elseif strcmp( z.B0030.cycle(i).type, 'discharge' ) 
%         
%         [reqStruct3] = z.B0030.cycle(i).data.Capacity;
%         C= transpose(reqStruct3);
%         xlswrite('Charge.xlsx',C,i+j,'C1');
%     else
%         
%     end
%     j=j+1;
%     k=k+1;
%     f=f+1;
% end
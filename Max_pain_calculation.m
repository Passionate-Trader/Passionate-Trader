clc
clear
close all

% tic
prev_data = xlsread('./Dataset/June Expiry 2021/15-30pm-31May21-option-chain-ED-NIFTY-03-Jun-2021.csv'); 
NSE_data = xlsread('./Dataset/June Expiry 2021/15-30pm-01Jun21-option-chain-ED-NIFTY-03-Jun-2021.csv');

days_to_expiry = 2;

[row,col] = size(NSE_data);
%% Max Pain Calculation
strike_price = NSE_data(:,11);
if size(NSE_data,1)>size(prev_data,1)
    for i=1:size(prev_data,1)
        pr_stk = prev_data(i,11);
        cu_stk = strike_price(i);
        if pr_stk~=cu_stk
            dummy_prev_data = [prev_data(1:i-1,:);zeros(1,size(prev_data,2));prev_data(i:end,:)];
            prev_data = dummy_prev_data;
        end
    end
elseif size(NSE_data,1)<size(prev_data,1)
    for i=1:size(NSE_data,1)
        cu_stk = NSE_data(i,11);
        pr_stk = strike_price(i);
        if cu_stk~=pr_stk
            dummy_NSE_data = [NSE_data(1:i-1,:);zeros(1,size(NSE_data,2));NSE_data(i:end,:)];
            NSE_data = dummy__data;
        end
    end
end

%%
call_OI = NSE_data(:,1);          call_OI(isnan(call_OI)) = 0;     % open interest
call_vol = NSE_data(:,3);         call_vol(isnan(call_vol)) = 0;   % volume
call_prm = NSE_data(:,5);         call_prm(isnan(call_prm)) = 0;   % premium value
call_iv = NSE_data(:,4);          call_iv(isnan(call_iv)) = 0;     % implied  volatility

put_OI = NSE_data(:,end);         put_OI(isnan(put_OI)) = 0;
put_vol = NSE_data(:,end-3+1);    put_vol(isnan(put_vol)) = 0;
put_prm = NSE_data(:,end-5+1);    put_prm(isnan(put_prm)) = 0;
put_iv = NSE_data(:,end-4+1);     put_iv(isnan(put_iv)) = 0;

%%
call_OI_prev = prev_data(:,1);       call_OI_prev(isnan(call_OI_prev)) = 0;
call_prm_prev = prev_data(:,5);      call_prm_prev(isnan(call_prm_prev)) = 0;
call_iv_prev = prev_data(:,4);       call_iv_prev(isnan(call_iv_prev)) = 0;

put_OI_prev = prev_data(:,end);      put_OI_prev(isnan(put_OI_prev)) = 0;
put_prm_prev = prev_data(:,end-5+1); put_prm_prev(isnan(put_prm_prev)) = 0;
put_iv_prev = prev_data(:,end-4+1);  put_iv_prev(isnan(put_iv_prev)) = 0;

%%
PCR_plot = put_OI./call_OI;
PCR_plot(PCR_plot>6.18) = 6.18;

for idx=1:length(strike_price)
    p = strike_price(idx);
    
    call_IV = max(p - strike_price,0);    % intrinsic value (spread between strike price & spot price)
    call_seller_loss(idx) = sum(call_IV.*call_OI);
    
    put_IV = max(strike_price - p,0);     % intrinsic value (spread between strike price & spot price)
    put_seller_loss(idx) = sum(put_IV.*put_OI);
end

pain = call_seller_loss + put_seller_loss;
[~,idx_stk] = min(pain);
max_pain = strike_price(idx_stk);


%% Historical Volatility count
nifty50_2020 = readtable('./Nifty50_historicalData/Yearly Data/2020.csv'); 
nifty50_2019 = readtable('./Nifty50_historicalData/Yearly Data/2019.csv'); 
nifty50_2018 = readtable('./Nifty50_historicalData/Yearly Data/2018.csv'); 
nifty50_2017 = readtable('./Nifty50_historicalData/Yearly Data/2017.csv'); 
nifty50_2016 = readtable('./Nifty50_historicalData/Yearly Data/2016.csv'); 

n_year = 3;    % Number of previous years to consider

if n_year == 3
    historical_NiftyData = cat(1,nifty50_2018,nifty50_2019,nifty50_2020);
elseif n_year == 5
    historical_NiftyData = cat(1,nifty50_2016,nifty50_2017,nifty50_2018,nifty50_2019,nifty50_2020);
elseif n_year == 1
    historical_NiftyData = nifty50_2020;
elseif n_year == 2
    historical_NiftyData = cat(1,nifty50_2019,nifty50_2020);
end
close_data = table2array(historical_NiftyData(:,5));
for k = 1:length(close_data)-1
    a = close_data(k+1);
    b = close_data(k);
    daily_returns(k,1) = log(a/b);
end
daily_volatility = std(daily_returns);

%%
margin = 2 * daily_volatility * sqrt(days_to_expiry);                  % Tolerance Margin
% margin = 0.05;                          
upper_strike = max_pain + max_pain*margin;      
lower_strike = max_pain - max_pain*margin;

range_lower = find(strike_price<lower_strike,1,'last');
range_upper = find(strike_price>upper_strike,1);

OI_PCR = sum(put_OI(range_lower:range_upper))/sum(call_OI(range_lower:range_upper));      % Open Interest Put-Call Ratio
Vol_PCR = sum(put_vol(range_lower:range_upper))/sum(call_vol(range_lower:range_upper));   % Volume Put-Call Ratio

% figure,
% bar(strike_price,call_seller_loss); hold on
% bar(strike_price,put_seller_loss); hold on
% plot(strike_price,pain,'color','black');
% legend("Call writer's loss","Put writer's loss","Net writer's Pain");
% title('Max Pain');

figure,
subplot(2,2,1);
yyaxis left
bar(strike_price(range_lower:range_upper),[call_OI(range_lower:range_upper),put_OI(range_lower:range_upper)]);
colororder('default');
% hold on
% bar(strike_price(range_lower:range_upper),[call_OI(range_lower:range_upper)]);
% hold off
title('Current Open Interest');
ylim([0,round(max(max(call_OI,put_OI)))]);
yyaxis right
plot(strike_price(range_lower:range_upper),PCR_plot(range_lower:range_upper),'black');
legend("Call OI","Put OI","PCR");

subplot(2,2,2);
bar(strike_price(range_lower:range_upper),[call_OI(range_lower:range_upper)-call_OI_prev(range_lower:range_upper),put_OI(range_lower:range_upper)-put_OI_prev(range_lower:range_upper)]);
legend("Call OI change","Put OI change"); title('Change in Open Interest');
% ylim([-round(max(max(abs(call_OI-call_OI_prev),abs(put_OI-put_OI_prev)))),round(max(max(abs(call_OI-call_OI_prev),abs(put_OI-put_OI_prev))))]);

subplot(2,2,3);
bar(strike_price(range_lower:range_upper),[call_prm(range_lower:range_upper)-call_prm_prev(range_lower:range_upper),put_prm(range_lower:range_upper)-put_prm_prev(range_lower:range_upper)]);
legend("Call prm change","Put prm change"); title('Change in premium value');
% ylim([-round(max(max(abs(call_prm-call_prm_prev),abs(put_prm-put_prm_prev)))),round(max(max(abs(call_prm-call_prm_prev),abs(put_prm-put_prm_prev))))]);

subplot(2,2,4);
bar(strike_price(range_lower:range_upper),[call_iv(range_lower:range_upper)-call_iv_prev(range_lower:range_upper),put_iv(range_lower:range_upper)-put_iv_prev(range_lower:range_upper)]);
legend("Call iv change","Put iv change"); title('Change in Implied Volatility');
% ylim([-round(max(max(abs(call_iv-call_iv_prev),abs(put_iv-put_iv_prev)))),round(max(max(abs(call_iv-call_iv_prev),abs(put_iv-put_iv_prev))))]);





fprintf('Open Interest Put-Call Ratio = %.4f \n',OI_PCR);
fprintf('Traded volume Put-Call Ratio = %.4f \n',Vol_PCR);
fprintf('Max Pain at price %.2f \n',max_pain);
% fprintf('Expected range of Nifty expiry for this data is [%.2f,%.2f] \n',lower_strike,upper_strike);

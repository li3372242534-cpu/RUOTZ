%% demo_kb_taper.m
%  KB microscope tapered-core (fang-xin) design helper.
%  Edit the CHANNELS table below to iterate on different design schemes.
%
%  For each channel you provide:
%     d1     small-end-to-target axial distance (mm)
%     theta  grazing angle of the mirror = lateral-face tilt (deg)
%     L      taper length along the optical axis (mm)
%  The script prints W1, W2, dW = W2-W1 and overlays the side-view
%  cross-sections of all tapers in one figure.

clear; clc;

%% ---------- 1. CHANNEL DESIGN TABLE ----------
% Each row = one taper / mirror group.
% Columns: { d1(mm), theta(deg), L(mm), label }
channels = { ...
    362.65, 0.5177, 12, 'M1&M2 (horizontal pair, near target)';
    383.39, 0.5512, 12, 'M3&M4 (vertical   pair, far  target)';
};

%% ---------- 2. COMPUTE ----------
n      = size(channels, 1);
d1_all = cell2mat(channels(:,1));
th_all = cell2mat(channels(:,2));
L_all  = cell2mat(channels(:,3));

[W1_all, W2_all] = kb_taper_widths(d1_all, th_all, L_all);
dW_all           = W2_all - W1_all;

%% ---------- 3. PRINT TABLE ----------
fprintf('\n');
fprintf('%-42s %10s %10s %8s %10s %10s %10s\n', ...
    'Channel', 'd1[mm]', 'theta[deg]', 'L[mm]', 'W1[mm]', 'W2[mm]', 'dW[mm]');
fprintf('%s\n', repmat('-', 1, 104));
for k = 1:n
    fprintf('%-42s %10.4f %10.4f %8.2f %10.4f %10.4f %10.4f\n', ...
        channels{k,4}, d1_all(k), th_all(k), L_all(k), ...
        W1_all(k), W2_all(k), dW_all(k));
end
fprintf('\n');

%% ---------- 4. PLOT (side-view cross sections) ----------
figure('Name', 'KB taper side view', 'Color', 'w');
hold on; grid on; box on;

cmap = lines(n);
for k = 1:n
    d1 = d1_all(k); L = L_all(k);
    W1 = W1_all(k); W2 = W2_all(k);

    % Trapezoid in the (axial, radial) plane.
    x = [d1,        d1+L,      d1+L,     d1,       d1];
    y = [+W1/2,    +W2/2,     -W2/2,    -W1/2,    +W1/2];

    plot(x, y, '-', 'Color', cmap(k,:), 'LineWidth', 1.6, ...
         'DisplayName', channels{k,4});

    % annotate W1 / W2
    text(d1,     +W1/2 + 0.5, sprintf('W1=%.4f', W1), ...
         'Color', cmap(k,:), 'HorizontalAlignment','center');
    text(d1+L,   +W2/2 + 0.5, sprintf('W2=%.4f', W2), ...
         'Color', cmap(k,:), 'HorizontalAlignment','center');
end

yline(0, 'k:', 'optical axis');
xlabel('Distance from target along optical axis  d  [mm]');
ylabel('Radial extent  [mm]');
title('KB tapered support cores -- side-view cross sections');
legend('Location', 'best');
axis tight;
ylim_now = ylim; ylim(ylim_now + [-1 1]);   % a bit of headroom for labels
hold off;

function [W1, W2] = kb_taper_widths(d1, theta_deg, L)
%KB_TAPER_WIDTHS  Compute the small-end and large-end widths of a KB
%                 microscope's tapered support core (fang-xin / cone-core).
%
%   [W1, W2] = KB_TAPER_WIDTHS(d1, theta_deg, L)
%
%   Inputs (each may be a scalar or a column vector; one element per channel):
%     d1         Axial distance from the target (source) to the SMALL end
%                of the taper, in millimeters.
%     theta_deg  Grazing angle of the mirror = tilt of each lateral face of
%                the taper relative to the optical axis, in degrees.
%     L          Axial length of the taper (along the optical axis), in mm.
%
%   Outputs:
%     W1   Edge length of the SQUARE cross-section at the small end (mm).
%     W2   Edge length of the SQUARE cross-section at the large end (mm).
%
%   Geometry (the taper is symmetric about the optical axis; all four
%   lateral faces are tilted outward by the same angle theta):
%
%       W1 = 2 * d1       * tan(theta)
%       W2 = 2 * (d1 + L) * tan(theta)
%          = W1 + 2 * L * tan(theta)
%
%   Example (reproduces M1&M2 row of the design table):
%       [w1, w2] = kb_taper_widths(362.65, 0.5177, 12)
%       % -> w1 ~ 6.5534,  w2 ~ 6.7703

    % --- input check ---
    if nargin < 3
        error('kb_taper_widths:NotEnoughInputs', ...
              'Need three inputs: d1 (mm), theta_deg (deg), L (mm).');
    end
    if any(theta_deg <= 0) || any(d1 < 0) || any(L <= 0)
        error('kb_taper_widths:BadValue', ...
              'd1 must be >= 0, theta_deg and L must be > 0.');
    end

    % --- core formula ---
    theta = theta_deg .* (pi/180);
    W1 = 2 .* d1       .* tan(theta);
    W2 = 2 .* (d1 + L) .* tan(theta);
end

function [match, cost] = FindCloestMatch(lower, higher, step, value)
    if value <= lower
        match = lower;
        cost = 10000;
        return
    end
    if value >= higher
        match = higher;
        cost = 100000;
        return
    end
    lookup = lower:step:higher;
    costs = 10000:10000:100000;
    lookup2 = lookup - value;
    lookup2(lookup2 < 0) = higher;
    [~, idx] = min(lookup2);
    match = lookup(idx);
    cost = costs(idx);
end 
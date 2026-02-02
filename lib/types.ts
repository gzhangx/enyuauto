const actions = ['校对','二校', '美编', '发布'] as const;
export type ActionType = typeof actions[number];
export interface UserProperty {
  readonly slackBotToken: string;
}

export const UserProperty = (slackBotToken: string): UserProperty => ({
  slackBotToken,
});

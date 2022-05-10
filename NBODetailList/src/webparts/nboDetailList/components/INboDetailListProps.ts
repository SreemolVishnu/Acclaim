import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INboDetailListProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
  sourceListName: string;
  nboStage: string;
  nboListName: string;
  classOfInsurance: string;
  industry: string;
  brokeragePercentage: string;
  teamList: string;
  pageSizeForPagination: number;
  emailNotificationSettings: string;
  complianceGroupEmail: string;
}

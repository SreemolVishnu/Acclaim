import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHrTrainingWpProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
  hrTrainingTitles: string;
  hrTrainingReports: string;
  TrainingModuleLibrary: string;
  webPartTitle: string;
  errorMessage: string;
  statementIfNoItems: string;
  labelForInstructions: string;
  messageBar: string;
}

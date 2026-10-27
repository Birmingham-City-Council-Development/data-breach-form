import { IBreachListItem } from '../../dataBreachForm/services/IBreachListItem';

export type FullItem = IBreachListItem & {
  Id: number;
  Author: { Title: string; EMail: string };
  Created: string;
  Data_x0020_Breach_x0020_ID: string;
  Status: string;
};

export type EditingState = Partial<IBreachListItem> & { Status?: string };

export type SortField = 'Created' | 'Status' | 'Data_x0020_Breach_x0020_ID' | 'Directorate_x0020_of_x0020_repor';

// Reads a field off the selected report as display text (HTML stripped).
export type FieldReader = (key: keyof FullItem) => string;

export type WorkbookSettings = {
  company_name: string;
  address: string;
  contact_number: string;
  logo_path: string;
  gate_pass_prefix: string;
  default_status: 'Returnable' | 'Non-Returnable';
};

export type WorkbookPass = {
  id: number;
  gate_pass_number: string;
  date: string;
  person_name: string;
  mobile_number: string;
  company_name: string;
  taken_by: string;
  authorized_by: string;
  status: 'Returnable' | 'Non-Returnable';
  items: Array<{
    id: number;
    description: string;
    pieces: number;
    feet: number;
    rate: number;
    amount: number;
  }>;
  total_items: number;
  total_pcs: number;
  total_feet: number;
  total_amount: number;
};

export type Batch = { toba_batchid: string; toba_name: string; toba_startdate?: string; toba_duedate?: string; toba_status?: string; };
export type Doc = { toba_documentid: string; toba_title: string; toba_version?: string; toba_requiressignature?: boolean; toba_fileurl?: string; toba_driveid?: string | null; toba_itemid?: string | null; toba_source?: 'sharepoint' | 'url' | null };

export type Business = {
	toba_businessid: string;
	toba_name: string;
	toba_code?: string;
	toba_isactive?: boolean;
	toba_description?: string;
};

export type BatchRecipient = {
	toba_batchrecipientid: string;
	toba_name: string;
	toba_User?: string;
	toba_Email?: string;
	toba_DisplayName?: string;
	toba_Department?: string;
	toba_JobTitle?: string;
	toba_Location?: string;
	toba_PrimaryGroup?: string;
	_toba_batch_value?: string; // lookup to batch
	_toba_business_value?: string; // lookup to business
};

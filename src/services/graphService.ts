import axios from 'axios';
export const fetchUserGroups = async (token: string): Promise<string[]> => {
  const res = await axios.get('https://graph.microsoft.com/v1.0/me/transitiveMemberOf?$select=displayName', { headers: { Authorization: `Bearer ${token}` } });
  return (res.data?.value || []).map((g: any) => g.displayName as string);
};

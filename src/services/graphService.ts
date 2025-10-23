export const fetchUserGroups = async (token: string): Promise<string[]> => {
  // Lazy-load axios to avoid ESM parsing issues in Jest without transforming node_modules
  const { default: axios } = await import('axios');
  const res = await axios.get('https://graph.microsoft.com/v1.0/me/transitiveMemberOf?$select=displayName', { headers: { Authorization: `Bearer ${token}` } });
  return (res.data?.value || []).map((g: any) => g.displayName as string);
};

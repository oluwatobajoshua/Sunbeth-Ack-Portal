  const saveBatch = async () => {
    try {
      if (!sqliteEnabled) return;
      const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');

      // Validate form data
      if (!batchForm.name.trim()) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Batch name is required' } }));
        return;
      }

      console.log('🚀 Starting batch creation process...', {
        batchName: batchForm.name,
        selectedDocs: batchForm.selectedDocuments.length,
        selectedUsers: batchForm.selectedUsers.length,
        selectedGroups: batchForm.selectedGroups.length,
        isEditing: !!editingBatchId
      });

      // Build recipients from selected users and expand selected groups into members
      const recipientSet = new Map<string, { address: string; name?: string }>();
      // Track origins (user/group) to apply business defaults
      const recipientOrigins = new Map<string, Set<string>>(); // emailLower -> Set of groupIds
      const addRecipient = (addrRaw: string, name?: string, originGroupId?: string) => {
        const addr = (addrRaw || '').trim();
        if (!addr) return;
        const key = addr.toLowerCase();
        if (!recipientSet.has(key)) recipientSet.set(key, { address: addr, name });
        if (originGroupId) {
          const set = recipientOrigins.get(key) || new Set<string>();
          set.add(originGroupId);
          recipientOrigins.set(key, set);
        }
      };

      for (const u of batchForm.selectedUsers) {
        addRecipient((u.mail || u.userPrincipalName || ''), u.displayName);
      }

      if (batchForm.selectedGroups.length > 0) {
        try {
          const token = await getGraphToken(['Group.Read.All', 'User.Read']);
          const membersArrays = await Promise.all(
            batchForm.selectedGroups.map(g => getGroupMembers(token, g.id).then(ms => ({ gid: g.id, members: ms })).catch(() => ({ gid: g.id, members: [] })))
          );
          for (const { gid, members } of membersArrays) {
            for (const m of members) {
              addRecipient((m.mail || m.userPrincipalName || ''), m.displayName, gid);
            }
          }
        } catch (e) {
          console.warn('Failed to expand group members for notifications', e);
        }
      }

      const recipients = Array.from(recipientSet.values());

      console.log('📊 Processed recipients:', {
        totalRecipients: recipients.length,
        sampleEmails: recipients.slice(0, 3).map(r => r.address)
      });

      // Helper maps for extra profile info
      const userByEmailLower = new Map<string, GraphUser>();
      for (const u of batchForm.selectedUsers) {
        const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
        if (email) userByEmailLower.set(email, u);
      }

      // Build email content (keeping this logic as is for notifications)
      const { subject, bodyHtml } = buildBatchEmail({
        appUrl: window.location.origin,
        batchName: batchForm.name,
        startDate: batchForm.startDate,
        dueDate: batchForm.dueDate,
        description: batchForm.description,
        documents: batchForm.selectedDocuments,
        senderDisplayName: user?.displayName || 'Sender'
      });

      // Determine who to notify (only new recipients if editing)
      const isNew = (addr: string) => !originalRecipientEmails.has(addr.trim().toLowerCase());
      let recipientsToNotify = recipients;
      if (editingBatchId && originalRecipientEmails.size > 0) {
        const filtered = recipients.filter(r => isNew(r.address));
        if (filtered.length === 0 && recipients.length > 0) recipientsToNotify = recipients;
        else recipientsToNotify = filtered;
      }
      if (batchForm.notifyByEmail && recipientsToNotify.length > 0) {
        await sendEmail(recipientsToNotify, subject, bodyHtml);
      }

      // ===== IMPROVED BATCH CREATION LOGIC =====
      let batchId: string | undefined;
      
      if (!editingBatchId) {
        // === STEP 1: CREATE BATCH ===
        console.log('📝 Step 1: Creating batch...');
        const createRes = await fetch(`${base}/api/batches`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: batchForm.name,
            startDate: batchForm.startDate || null,
            dueDate: batchForm.dueDate || null,
            description: batchForm.description || null,
            status: 1
          })
        });
        
        if (!createRes.ok) {
          const errorText = await createRes.text().catch(() => '');
          console.error('❌ Batch creation failed:', createRes.status, errorText);
          throw new Error(`Batch creation failed: ${createRes.status} - ${errorText}`);
        }
        
        const createJson = await createRes.json();
        const batchIdRaw = (createJson?.id ?? createJson?.batchId ?? createJson?.toba_batchid ?? createJson?.ID);
        batchId = typeof batchIdRaw === 'string' ? batchIdRaw : (Number.isFinite(Number(batchIdRaw)) ? String(batchIdRaw) : undefined);
        
        console.log('✅ Batch created successfully:', {
          createJson,
          batchIdRaw,
          finalBatchId: batchId
        });
        
        if (!batchId) {
          throw new Error('Batch ID is missing from server response');
        }

        // Small delay to ensure batch is fully committed
        await new Promise(resolve => setTimeout(resolve, 150));

        // === STEP 2: VERIFY BATCH EXISTS ===
        console.log('🔍 Step 2: Verifying batch exists...');
        const verifyRes = await fetch(`${base}/api/batches`);
        const allBatches = await verifyRes.json();
        const createdBatch = allBatches.find((b: any) => String(b.id) === String(batchId));
        
        if (!createdBatch) {
          console.error('❌ Batch verification failed - batch not found after creation');
          throw new Error('Batch verification failed - batch not found');
        }
        
        console.log('✅ Batch verified:', createdBatch.name);

      } else {
        // EDITING MODE
        batchId = editingBatchId;
        console.log('📝 Updating existing batch:', batchId);
        
        const updateRes = await fetch(`${base}/api/batches/${encodeURIComponent(batchId)}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: batchForm.name,
            startDate: batchForm.startDate || null,
            dueDate: batchForm.dueDate || null,
            description: batchForm.description || null,
            status: 1
          })
        });
        
        if (!updateRes.ok) {
          const errorText = await updateRes.text().catch(() => '');
          throw new Error(`Batch update failed: ${updateRes.status} - ${errorText}`);
        }
        
        console.log('✅ Batch updated successfully');
      }

      // === STEP 3: CREATE DOCUMENTS ===
      const allDocsPayload = batchForm.selectedDocuments.map(d => ({
        title: d.title,
        url: d.url,
        version: d.version ?? 1,
        requiresSignature: !!d.requiresSignature,
        driveId: (d as any).driveId || null,
        itemId: (d as any).itemId || null,
        source: (d as any).source || null
      }));
      
      const docsToPost = !editingBatchId
        ? allDocsPayload
        : allDocsPayload.filter(d => !originalDocUrls.has((d.url || '').trim()));
      
      console.log('📄 Step 3: Processing documents...', {
        isCreating: !editingBatchId,
        totalDocs: allDocsPayload.length,
        docsToPost: docsToPost.length,
        batchId
      });
      
      if (docsToPost.length > 0) {
        console.log('📄 Creating documents for batch:', batchId);
        const docsRes = await fetch(`${base}/api/batches/${batchId}/documents`, {
          method: 'POST', 
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ documents: docsToPost })
        });
        
        if (!docsRes.ok) {
          const errorText = await docsRes.text().catch(() => '');
          console.error('❌ Documents insert failed:', docsRes.status, errorText);
          throw new Error(`Documents creation failed: ${docsRes.status} - ${errorText}`);
        } else {
          const docsResult = await docsRes.json().catch(() => null);
          console.log('✅ Documents created successfully:', docsResult);
        }
      } else {
        console.log('⏭️ No documents to create');
      }

      // === STEP 4: CREATE RECIPIENTS ===
      const recipientsPayloadAll = recipients.map(r => {
        const emailLower = (r.address || '').toLowerCase();
        const u = userByEmailLower.get(emailLower);
        let primaryGroupName: string | undefined = undefined;
        const origins = recipientOrigins.get(emailLower);
        if (origins && origins.size > 0) {
          const firstGid = origins.values().next().value as string;
          const g = batchForm.selectedGroups.find(x => x.id === firstGid);
          if (g?.displayName) primaryGroupName = g.displayName;
        }
        const mappedBusinessId = (businessMap[emailLower] ?? (defaultBusinessId !== '' ? Number(defaultBusinessId) : null));
        return {
          businessId: mappedBusinessId,
          user: emailLower,
          email: emailLower,
          userEmail: emailLower,
          userPrincipalName: emailLower,
          displayName: r.name || undefined,
          department: u?.department || undefined,
          jobTitle: u?.jobTitle || undefined,
          location: u?.officeLocation || undefined,
          primaryGroup: primaryGroupName || undefined
        };
      });
      
      const recipientsPayload = editingBatchId
        ? recipientsPayloadAll.filter(r => !originalRecipientEmails.has((r.email || '').trim().toLowerCase()))
        : recipientsPayloadAll;
      
      console.log('👥 Step 4: Processing recipients...', {
        isCreating: !editingBatchId,
        totalRecipients: recipientsPayloadAll.length,
        recipientsToPost: recipientsPayload.length,
        batchId,
        sampleRecipient: recipientsPayload[0]
      });
      
      if (recipientsPayload.length > 0) {
        console.log('👥 Creating recipients for batch:', batchId);
        const recRes = await fetch(`${base}/api/batches/${batchId}/recipients`, {
          method: 'POST', 
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ recipients: recipientsPayload })
        });
        
        if (!recRes.ok) {
          const errorText = await recRes.text().catch(() => '');
          console.error('❌ Recipients insert failed:', recRes.status, errorText);
          throw new Error(`Recipients creation failed: ${recRes.status} - ${errorText}`);
        } else {
          const recipientsResult = await recRes.json().catch(() => null);
          console.log('✅ Recipients created successfully:', recipientsResult);
        }

        // === STEP 5: VERIFY RELATIONSHIPS ===
        console.log('🔍 Step 5: Verifying relationships...');
        try {
          const verify = await fetch(`${base}/api/batches/${batchId}/recipients`, { cache: 'no-store' });
          const rows = await verify.json();
          if (!Array.isArray(rows) || rows.length === 0) {
            console.warn('⚠️ Recipients verification returned empty for batch', batchId);
            window.dispatchEvent(new CustomEvent('sunbeth:toast', { 
              detail: { message: `Warning: ${recipientsPayload.length} recipients were sent but verification shows none linked to batch` } 
            }));
          } else {
            console.log('✅ Recipients verification successful:', rows.length, 'recipients linked to batch');
          }
        } catch (e) {
          console.warn('⚠️ Recipients verification failed', e);
        }
      } else {
        console.log('⏭️ No recipients to create');
      }

      // === FINAL SUCCESS ===
      const actionWord = editingBatchId ? 'updated' : 'created';
      const summary = [
        `Batch "${batchForm.name}" ${actionWord}`,
        docsToPost.length > 0 ? `${docsToPost.length} documents added` : '',
        recipientsPayload.length > 0 ? `${recipientsPayload.length} recipients added` : '',
        batchForm.notifyByEmail ? 'email notifications sent' : ''
      ].filter(Boolean).join(', ');

      console.log('🎉 Batch creation completed successfully:', summary);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: summary } }));

      // Reset form
      setBatchForm({
        name: '',
        startDate: '',
        dueDate: '',
        description: '',
        selectedUsers: [],
        selectedGroups: [],
        selectedDocuments: [],
        notifyByEmail: true,
        notifyByTeams: false
      });
      setBusinessMap({});
      setDefaultBusinessId('');
      setEditingBatchId(null);
      setOriginalRecipientEmails(new Set());
      setOriginalDocUrls(new Set());

    } catch (e: any) {
      console.error('❌ Batch creation failed:', e);
      const errorMessage = e?.message || 'Unknown error occurred';
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { 
        detail: { message: `Failed to create batch: ${errorMessage}` } 
      }));
    }
  };
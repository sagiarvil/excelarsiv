#!/usr/bin/env node
import { execSync } from 'node:child_process';

async function cleanup() {
  try {
    const token = execSync('gcloud auth print-access-token', { encoding: 'utf8' }).trim();
    const site = process.env.FIREBASE_HOSTING_SITE || 'excelarsiv';
    const project = process.env.FIREBASE_PROJECT_ID || 'carbon-web-1265b';
    
    // 1. Get current active release to protect active version
    const releasesRes = await fetch(`https://firebasehosting.googleapis.com/v1beta1/projects/${project}/sites/${site}/releases?pageSize=5`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const releasesData = await releasesRes.json();
    const activeVersionName = releasesData?.releases?.[0]?.version?.name;
    console.log(`Active release version to protect: ${activeVersionName}`);

    // 2. Fetch versions and delete older versions
    let pageToken = '';
    let deletedCount = 0;
    do {
      const url = `https://firebasehosting.googleapis.com/v1beta1/projects/${project}/sites/${site}/versions?pageSize=50${pageToken ? `&pageToken=${pageToken}` : ''}`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      const data = await res.json();
      const versions = data.versions || [];
      for (const v of versions) {
        if (v.name && v.name !== activeVersionName && v.status !== 'DELETED') {
          console.log(`Deleting old hosting version ${v.name} (${v.status})...`);
          const delRes = await fetch(`https://firebasehosting.googleapis.com/v1beta1/${v.name}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${token}` }
          });
          if (delRes.ok) deletedCount++;
        }
      }
      pageToken = data.nextPageToken || '';
    } while (pageToken);
    console.log(`Hosting storage cleanup completed: ${deletedCount} versions deleted.`);
  } catch (err) {
    console.warn('Storage cleanup warning:', err.message);
  }
}

cleanup();

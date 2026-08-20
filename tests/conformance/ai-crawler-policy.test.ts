import { describe, it } from 'node:test';
import assert from 'node:assert';
import fs from 'node:fs';

describe('AI Crawler Policy Contract', () => {
  it('robots.txt must allow OAI-SearchBot and ChatGPT-User', () => {
    const robots = fs.readFileSync('public/robots.txt', 'utf8');
    assert.ok(robots.includes('User-agent: OAI-SearchBot'), 'Missing OAI-SearchBot');
    assert.ok(robots.includes('User-agent: ChatGPT-User'), 'Missing ChatGPT-User');
  });
});

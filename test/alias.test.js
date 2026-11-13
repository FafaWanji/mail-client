import test from 'node:test';
import assert from 'node:assert/strict';
import { extractAlias, aliasAddress, senderDomain } from '../src/alias.js';

const me = 'me@gmail.com';
test('plain', () => assert.equal(extractAlias('me+spotify@gmail.com', me), 'spotify'));
test('display name and case', () => assert.equal(extractAlias('"X" <Me+Spotify@Gmail.com>', me), 'spotify'));
test('several recipients', () => assert.equal(extractAlias('a@x.com, me+uber@gmail.com', me), 'uber'));
test('gmail dots ignored, tag kept', () => assert.equal(extractAlias('m.e+net-flix.1@gmail.com', me), 'net-flix.1'));
test('stranger plus address', () => assert.equal(extractAlias('other+tag@gmail.com', me), null));
test('no tag', () => assert.equal(extractAlias('me@gmail.com', me), null));
test('empty tag', () => assert.equal(extractAlias('me+@gmail.com', me), null));
test('empty header', () => assert.equal(extractAlias('', me), null));
test('aliasAddress', () => assert.equal(aliasAddress(me, 'spotify'), 'me+spotify@gmail.com'));
test('senderDomain', () => assert.equal(senderDomain('Spotify <no-reply@email.spotify.com>'), 'spotify.com'));
test('senderDomain none', () => assert.equal(senderDomain('garbage'), null));

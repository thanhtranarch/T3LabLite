# -*- coding: utf-8 -*-
"""Deployed MCP bridge kept current (lib/bridge_deploy.py): MCP Control
deploys and takes ownership, startup.py refreshes only what this extension
owns. Runs on Linux, no Revit."""
import os
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = REPO
LIB = os.path.join(EXT, 'lib')
sys.path.insert(0, LIB)

import bridge_deploy as D  # noqa: E402


def write(path, text):
    folder = os.path.dirname(path)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


class DeployTests(unittest.TestCase):
    def setUp(self):
        root = tempfile.mkdtemp(prefix='t3lab_deploy_')
        self.target = os.path.join(root, 'AppData', 'T3LabAI')
        self.dev = os.path.join(root, 'dev', 'lib', 'core')
        self.lite = os.path.join(root, 'lite', 'lib', 'core')
        for core, tag in ((self.dev, 'dev'), (self.lite, 'lite')):
            write(os.path.join(core, 'bridge.py'), '# bridge {}\n'.format(tag))
            write(os.path.join(core, 'hub_store.py'), '# hub {}\n'.format(tag))
            write(os.path.join(core, 'skills_registry.json'), '{}')

    def test_deploy_copies_companions_and_takes_ownership(self):
        path, error, written = D.deploy(self.target, self.dev)
        self.assertIsNone(error)
        self.assertEqual(path, os.path.join(self.target, 'bridge.py'))
        self.assertEqual(sorted(written), ['bridge.py', 'hub_store.py', 'skills.json'])
        self.assertEqual(read(os.path.join(self.target, 'hub_store.py')), '# hub dev\n')
        self.assertFalse(os.path.exists(os.path.join(self.target, 'mcp_providers.py')))
        self.assertTrue(D._same_dir(D.owner(self.target), self.dev))
        self.assertEqual(D.deploy(self.target, self.dev)[2], [])      # unchanged: no write

    def test_missing_bridge_is_an_error(self):
        os.remove(os.path.join(self.dev, 'bridge.py'))
        path, error, _ = D.deploy(self.target, self.dev)
        self.assertIsNone(path)
        self.assertIn('bridge.py', error)

    def test_refresh_never_deploys_a_first_time(self):
        self.assertEqual(D.refresh(self.target, self.dev)[0], 'skipped')
        self.assertFalse(os.path.exists(os.path.join(self.target, 'bridge.py')))

    def test_refresh_updates_what_this_extension_owns(self):
        D.deploy(self.target, self.dev)
        self.assertEqual(D.refresh(self.target, self.dev), ('current', ''))
        write(os.path.join(self.dev, 'hub_store.py'), '# hub dev v2\n')
        self.assertEqual(D.refresh(self.target, self.dev), ('updated', 'hub_store.py'))
        self.assertEqual(read(os.path.join(self.target, 'hub_store.py')), '# hub dev v2\n')

    def test_refresh_leaves_another_extensions_bridge_alone(self):
        D.deploy(self.target, self.lite)                  # Lite's MCP Control connected last
        state, detail = D.refresh(self.target, self.dev)
        self.assertEqual(state, 'skipped')
        self.assertIn('another extension', detail)
        self.assertEqual(read(os.path.join(self.target, 'bridge.py')), '# bridge lite\n')
        D.deploy(self.target, self.dev)                   # dev's MCP Control takes it back
        self.assertEqual(D.refresh(self.target, self.lite)[0], 'skipped')
        self.assertEqual(read(os.path.join(self.target, 'bridge.py')), '# bridge dev\n')

    def test_bridge_from_before_ownership_is_adopted(self):
        write(os.path.join(self.target, 'bridge.py'), '# old bridge\n')
        self.assertEqual(D.refresh(self.target, self.dev)[0], 'updated')
        self.assertEqual(read(os.path.join(self.target, 'bridge.py')), '# bridge dev\n')
        self.assertTrue(D._same_dir(D.owner(self.target), self.dev))

    def test_hub_store_ships_with_lite(self):
        # Lite has no skills / providers companions: deploy() skips missing ones.
        self.assertTrue(os.path.isfile(os.path.join(D.core_dir(), 'hub_store.py')))
        self.assertIn(('hub_store.py', 'hub_store.py'), D.COMPANIONS)


class WiringTests(unittest.TestCase):
    def test_startup_refreshes_before_the_engine_checks(self):
        src = read(os.path.join(EXT, 'startup.py'))
        main = src.split('def main():', 1)[1]
        self.assertIn('refresh_mcp_bridge()', main)
        self.assertLess(main.index('refresh_mcp_bridge()'), main.index('pythonnet_half_started()'))
        self.assertIn('_load_lib_module("bridge_deploy")', src)

    def test_mcp_control_deploys_through_the_same_module(self):
        src = read(os.path.join(LIB, 'Services', 'mcp_service.py'))
        body = src.split('def deploy_bridge():', 1)[1].split('@staticmethod', 1)[0]
        self.assertIn('bridge_deploy.deploy(', body)


if __name__ == '__main__':
    unittest.main()
